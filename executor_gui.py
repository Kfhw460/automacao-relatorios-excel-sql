"""
Executor visual da automação de relatório.

Abre uma janela amigável para quem não conhece Python:
  - Botão grande para iniciar a automação
  - Etapas com indicação visual (pendente / em andamento / concluída / erro)
  - Log em tempo real, colorido
  - Cronômetro com o tempo total de execução
  - Resultado final claro (sucesso ou falha)

Como abrir:
  - Dê dois cliques em "Executar Relatorio.bat", ou
  - Rode:  pythonw executor_gui.py
"""

import contextlib
import importlib
import os
import queue
import sys
import threading
import time
import traceback
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import ttk, messagebox

if getattr(sys, "frozen", False):
    PASTA_APP = Path(sys.executable).parent
else:
    PASTA_APP = Path(__file__).resolve().parent

os.chdir(PASTA_APP)
if str(PASTA_APP) not in sys.path:
    sys.path.insert(0, str(PASTA_APP))

COR_FUNDO = "#f4f6fa"
COR_CARD = "#ffffff"
COR_PRIMARIA = "#1f6feb"
COR_SUCESSO = "#1a7f37"
COR_ERRO = "#cf222e"
COR_TEXTO = "#1f2328"
COR_TEXTO_SUAVE = "#656d76"
COR_LOG_FUNDO = "#0d1117"
COR_LOG_TEXTO = "#e6edf3"

ETAPAS = [
    ("config", "Carregar configurações (.env)", None),
    ("backup", "Criar cópia de segurança (backup)", "Cópia de Segurança"),
    ("copia", "Criar cópia para atualização", "Cópia de Produção"),
    ("excel", "Atualizar dados do Excel (banco de dados)", "Sucesso: Dados atualizados"),
]

ICONES = {
    "pendente": ("○", COR_TEXTO_SUAVE),
    "andamento": ("◐", COR_PRIMARIA),
    "ok": ("✔", COR_SUCESSO),
    "erro": ("✖", COR_ERRO),
}


def formatar_tempo(segundos: float) -> str:
    h, resto = divmod(int(segundos), 3600)
    m, s = divmod(resto, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def formatar_tempo_extenso(segundos: float) -> str:
    if segundos < 60:
        return f"{segundos:.1f} segundos"
    m, s = divmod(int(segundos), 60)
    if m < 60:
        return f"{m} min e {s} seg"
    h, m = divmod(m, 60)
    return f"{h} h, {m} min e {s} seg"


class _EscritorFila:
    """Redireciona print() da automação para a fila da interface."""

    def __init__(self, fila: queue.Queue):
        self.fila = fila
        self._buffer = ""

    def write(self, texto: str) -> int:
        self._buffer += texto
        while "\n" in self._buffer:
            linha, self._buffer = self._buffer.split("\n", 1)
            if linha.strip():
                self.fila.put(("print", linha))
        return len(texto)

    def flush(self) -> None:
        if self._buffer.strip():
            self.fila.put(("print", self._buffer))
        self._buffer = ""


_modulo_automacao = None


def executar_em_segundo_plano(fila: queue.Queue) -> None:
    global _modulo_automacao

    try:
        import pythoncom
        pythoncom.CoInitialize()
    except ImportError:
        pythoncom = None

    escritor = _EscritorFila(fila)
    try:
        with contextlib.redirect_stdout(escritor), contextlib.redirect_stderr(escritor):
            fila.put(("etapa", "config", "andamento"))
            try:
                if _modulo_automacao is None:
                    import automacao_relatorio
                    _modulo_automacao = automacao_relatorio
                else:
                    _modulo_automacao = importlib.reload(_modulo_automacao)
            except Exception as e:
                fila.put(("etapa", "config", "erro"))
                fila.put(("log", "erro", f"Não foi possível carregar as configurações: {e}"))
                fila.put(("fim", False, str(e)))
                return

            mod = _modulo_automacao
            fila.put(("etapa", "config", "ok"))
            fila.put(("arquivo_log", str(mod.ARQUIVO_LOG)))

            registrar_original = mod.registrar_log
            estado = {"sucesso": False, "erro": None}

            def registrar_log_visual(mensagem: str) -> None:
                registrar_original(mensagem)
                texto = mensagem.lower()
                if texto.startswith("erro"):
                    nivel = "erro"
                    estado["erro"] = mensagem
                elif texto.startswith("aviso"):
                    nivel = "aviso"
                elif texto.startswith("sucesso") or "100% de sucesso" in texto:
                    nivel = "sucesso"
                else:
                    nivel = "info"
                fila.put(("log", nivel, mensagem))

                for chave, _, gatilho in ETAPAS:
                    if gatilho and gatilho in mensagem:
                        fila.put(("etapa", chave, "ok"))
                if "Cópia de Segurança" in mensagem:
                    fila.put(("etapa", "copia", "andamento"))
                elif "Cópia de Produção" in mensagem:
                    fila.put(("etapa", "excel", "andamento"))
                    fila.put(("log", "info", "Atualizando o Excel... isso pode levar alguns minutos."))
                if "100% de sucesso" in mensagem:
                    estado["sucesso"] = True

            mod.registrar_log = registrar_log_visual
            try:
                fila.put(("etapa", "backup", "andamento"))
                mod.executar_automacao_relatorio()
            finally:
                mod.registrar_log = registrar_original

            escritor.flush()
            if estado["sucesso"]:
                fila.put(("fim", True, None))
            else:
                fila.put(("fim", False, estado["erro"] or "A automação não foi concluída."))

    except Exception as e:
        fila.put(("log", "erro", f"Erro inesperado: {e}"))
        fila.put(("log", "erro", traceback.format_exc()))
        fila.put(("fim", False, str(e)))
    finally:
        if pythoncom is not None:
            pythoncom.CoUninitialize()


class ExecutorApp:
    def __init__(self, raiz: tk.Tk):
        self.raiz = raiz
        self.fila: queue.Queue = queue.Queue()
        self.executando = False
        self.inicio = 0.0
        self.caminho_log = None

        raiz.title("Automação de Relatório - Executor")
        raiz.geometry("880x700")
        raiz.minsize(720, 560)
        raiz.configure(bg=COR_FUNDO)
        raiz.protocol("WM_DELETE_WINDOW", self.ao_fechar)

        self._configurar_estilos()
        self._montar_tela()
        self.raiz.after(100, self._processar_fila)

    def _configurar_estilos(self) -> None:
        estilo = ttk.Style()
        try:
            estilo.theme_use("clam")
        except tk.TclError:
            pass
        estilo.configure("Fundo.TFrame", background=COR_FUNDO)
        estilo.configure(
            "Executar.TButton", font=("Segoe UI", 13, "bold"), padding=(24, 12),
            foreground="white", background=COR_PRIMARIA, borderwidth=0,
        )
        estilo.map(
            "Executar.TButton",
            background=[("disabled", "#8bb3f0"), ("active", "#1858c4")],
            foreground=[("disabled", "white")],
        )
        estilo.configure("Secundario.TButton", font=("Segoe UI", 10), padding=(12, 6))
        estilo.configure(
            "Barra.Horizontal.TProgressbar", troughcolor="#e1e6ee",
            background=COR_PRIMARIA, thickness=8, borderwidth=0,
        )

    def _montar_tela(self) -> None:
        cab = tk.Frame(self.raiz, bg=COR_PRIMARIA)
        cab.pack(fill="x")
        tk.Label(
            cab, text="📊  Automação de Relatório Excel", bg=COR_PRIMARIA, fg="white",
            font=("Segoe UI", 18, "bold"),
        ).pack(side="left", padx=20, pady=14)

        corpo = ttk.Frame(self.raiz, style="Fundo.TFrame", padding=16)
        corpo.pack(fill="both", expand=True)

        card = tk.Frame(corpo, bg=COR_CARD, highlightbackground="#d0d7de", highlightthickness=1)
        card.pack(fill="x")

        esquerda = tk.Frame(card, bg=COR_CARD)
        esquerda.pack(side="left", fill="both", expand=True, padx=18, pady=16)

        self.lbl_status = tk.Label(
            esquerda, text="⏸  Pronto para executar", bg=COR_CARD, fg=COR_TEXTO,
            font=("Segoe UI", 15, "bold"), anchor="w",
        )
        self.lbl_status.pack(fill="x")
        self.lbl_detalhe = tk.Label(
            esquerda, text="Clique em  ▶ Executar  para iniciar a atualização do relatório.",
            bg=COR_CARD, fg=COR_TEXTO_SUAVE, font=("Segoe UI", 10), anchor="w",
            justify="left", wraplength=480,
        )
        self.lbl_detalhe.pack(fill="x", pady=(4, 10))

        self.barra = ttk.Progressbar(esquerda, mode="indeterminate", style="Barra.Horizontal.TProgressbar")
        self.barra.pack(fill="x")

        direita = tk.Frame(card, bg=COR_CARD)
        direita.pack(side="right", padx=18, pady=16)

        tk.Label(direita, text="Tempo de execução", bg=COR_CARD, fg=COR_TEXTO_SUAVE,
                 font=("Segoe UI", 9)).pack()
        self.lbl_tempo = tk.Label(direita, text="00:00:00", bg=COR_CARD, fg=COR_TEXTO,
                                  font=("Consolas", 26, "bold"))
        self.lbl_tempo.pack(pady=(0, 8))
        self.btn_executar = ttk.Button(direita, text="▶  Executar", style="Executar.TButton",
                                       command=self.iniciar, cursor="hand2")
        self.btn_executar.pack()

        card_etapas = tk.Frame(corpo, bg=COR_CARD, highlightbackground="#d0d7de", highlightthickness=1)
        card_etapas.pack(fill="x", pady=(12, 0))
        tk.Label(card_etapas, text="Etapas", bg=COR_CARD, fg=COR_TEXTO,
                 font=("Segoe UI", 11, "bold"), anchor="w").pack(fill="x", padx=18, pady=(10, 4))

        self.lbl_etapas = {}
        for chave, texto, _ in ETAPAS:
            linha = tk.Frame(card_etapas, bg=COR_CARD)
            linha.pack(fill="x", padx=18, pady=2)
            icone = tk.Label(linha, bg=COR_CARD, font=("Segoe UI", 12, "bold"), width=2)
            icone.pack(side="left")
            rotulo = tk.Label(linha, text=texto, bg=COR_CARD, font=("Segoe UI", 10), anchor="w")
            rotulo.pack(side="left", fill="x")
            self.lbl_etapas[chave] = (icone, rotulo)
        tk.Frame(card_etapas, bg=COR_CARD, height=8).pack()
        self._resetar_etapas()

        topo_log = ttk.Frame(corpo, style="Fundo.TFrame")
        topo_log.pack(fill="x", pady=(14, 4))
        tk.Label(topo_log, text="Log da execução", bg=COR_FUNDO, fg=COR_TEXTO,
                 font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Button(topo_log, text="Limpar", style="Secundario.TButton",
                   command=self.limpar_log).pack(side="right")
        self.btn_abrir_log = ttk.Button(topo_log, text="Abrir log completo", style="Secundario.TButton",
                                        command=self.abrir_log, state="disabled")
        self.btn_abrir_log.pack(side="right", padx=(0, 6))
        self.btn_abrir_pasta = ttk.Button(topo_log, text="Abrir pasta dos relatórios",
                                          style="Secundario.TButton", command=self.abrir_pasta,
                                          state="disabled")
        self.btn_abrir_pasta.pack(side="right", padx=(0, 6))

        quadro_log = tk.Frame(corpo, bg=COR_LOG_FUNDO)
        quadro_log.pack(fill="both", expand=True)
        self.txt_log = tk.Text(
            quadro_log, bg=COR_LOG_FUNDO, fg=COR_LOG_TEXTO, font=("Consolas", 10),
            relief="flat", wrap="word", padx=10, pady=8, state="disabled",
        )
        rolagem = ttk.Scrollbar(quadro_log, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=rolagem.set)
        rolagem.pack(side="right", fill="y")
        self.txt_log.pack(side="left", fill="both", expand=True)

        self.txt_log.tag_configure("hora", foreground="#7d8590")
        self.txt_log.tag_configure("info", foreground=COR_LOG_TEXTO)
        self.txt_log.tag_configure("sucesso", foreground="#3fb950")
        self.txt_log.tag_configure("aviso", foreground="#d29922")
        self.txt_log.tag_configure("erro", foreground="#f85149")
        self.txt_log.tag_configure("destaque", foreground="#79c0ff", font=("Consolas", 10, "bold"))

        self.lbl_rodape = tk.Label(self.raiz, text="", bg=COR_FUNDO, fg=COR_TEXTO_SUAVE,
                                   font=("Segoe UI", 9), anchor="w")
        self.lbl_rodape.pack(fill="x", padx=16, pady=(0, 8))

    def iniciar(self) -> None:
        if self.executando:
            return
        self.executando = True
        self.inicio = time.time()
        self._resetar_etapas()
        self.btn_executar.configure(state="disabled", text="⏳  Executando...")
        self._definir_status("⏳  Executando...", COR_PRIMARIA,
                             "A automação está em andamento. Não feche esta janela nem o Excel.")
        self.barra.start(12)
        self.escrever_log("destaque", "=" * 60)
        self.escrever_log("destaque", f"Nova execução iniciada em {datetime.now():%d/%m/%Y às %H:%M:%S}")
        self.escrever_log("destaque", "=" * 60)

        threading.Thread(target=executar_em_segundo_plano, args=(self.fila,), daemon=True).start()
        self._atualizar_cronometro()

    def finalizar(self, sucesso: bool, erro) -> None:
        self.executando = False
        duracao = time.time() - self.inicio
        self.lbl_tempo.configure(text=formatar_tempo(duracao))
        self.barra.stop()
        self.btn_executar.configure(state="normal", text="▶  Executar novamente")
        tempo_txt = formatar_tempo_extenso(duracao)

        if sucesso:
            self._definir_status("✔  Concluído com sucesso!", COR_SUCESSO,
                                 f"O relatório foi atualizado em {tempo_txt}.")
            self.escrever_log("sucesso", f"✔ FINALIZADO COM SUCESSO — tempo total: {tempo_txt}")
        else:
            for chave, (icone, _) in self.lbl_etapas.items():
                if icone.cget("text") == ICONES["andamento"][0]:
                    self._definir_etapa(chave, "erro")
            self._definir_status("✖  A execução falhou", COR_ERRO,
                                 f"Problema: {erro}\nTempo até a falha: {tempo_txt}. Veja o log abaixo.")
            self.escrever_log("erro", f"✖ FALHOU — tempo até a falha: {tempo_txt}")

        self.lbl_rodape.configure(text=f"Última execução: {datetime.now():%d/%m/%Y %H:%M:%S}  •  "
                                       f"{'Sucesso' if sucesso else 'Falha'}  •  Duração: {tempo_txt}")
        self.raiz.bell()

    def limpar_log(self) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    def abrir_log(self) -> None:
        if self.caminho_log and Path(self.caminho_log).exists():
            os.startfile(self.caminho_log)
        else:
            messagebox.showinfo("Log", "O arquivo de log ainda não foi criado.")

    def abrir_pasta(self) -> None:
        if self.caminho_log and Path(self.caminho_log).parent.exists():
            os.startfile(Path(self.caminho_log).parent)
        else:
            messagebox.showinfo("Pasta", "A pasta dos relatórios ainda não existe.")

    def ao_fechar(self) -> None:
        if self.executando and not messagebox.askyesno(
            "Execução em andamento",
            "A automação ainda está rodando.\nFechar agora pode deixar o Excel aberto em segundo plano.\n\n"
            "Deseja fechar mesmo assim?",
        ):
            return
        self.raiz.destroy()

    def escrever_log(self, nivel: str, mensagem: str) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", f"[{datetime.now():%H:%M:%S}] ", "hora")
        self.txt_log.insert("end", mensagem.rstrip() + "\n", nivel)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _definir_status(self, titulo: str, cor: str, detalhe: str) -> None:
        self.lbl_status.configure(text=titulo, fg=cor)
        self.lbl_detalhe.configure(text=detalhe)

    def _definir_etapa(self, chave: str, estado: str) -> None:
        icone, rotulo = self.lbl_etapas[chave]
        simbolo, cor = ICONES[estado]
        icone.configure(text=simbolo, fg=cor)
        rotulo.configure(fg=COR_TEXTO if estado != "pendente" else COR_TEXTO_SUAVE)

    def _resetar_etapas(self) -> None:
        for chave in self.lbl_etapas:
            self._definir_etapa(chave, "pendente")

    def _atualizar_cronometro(self) -> None:
        if self.executando:
            self.lbl_tempo.configure(text=formatar_tempo(time.time() - self.inicio))
            self.raiz.after(500, self._atualizar_cronometro)

    def _processar_fila(self) -> None:
        try:
            while True:
                item = self.fila.get_nowait()
                tipo = item[0]
                if tipo == "log":
                    self.escrever_log(item[1], item[2])
                elif tipo == "print":
                    texto = item[1]
                    if "[ERRO" in texto or "Traceback" in texto:
                        nivel = "erro"
                    elif "sucesso" in texto.lower():
                        nivel = "sucesso"
                    else:
                        nivel = "info"
                    self.escrever_log(nivel, texto)
                elif tipo == "etapa":
                    self._definir_etapa(item[1], item[2])
                elif tipo == "arquivo_log":
                    self.caminho_log = item[1]
                    self.btn_abrir_log.configure(state="normal")
                    self.btn_abrir_pasta.configure(state="normal")
                elif tipo == "fim":
                    self.finalizar(item[1], item[2])
        except queue.Empty:
            pass
        self.raiz.after(100, self._processar_fila)


def main() -> None:
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    raiz = tk.Tk()
    ExecutorApp(raiz)
    raiz.mainloop()


if __name__ == "__main__":
    main()
