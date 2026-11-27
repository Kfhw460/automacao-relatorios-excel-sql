from datetime import datetime
from pathlib import Path
import os
import shutil
import time
import xlwings as xw

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


def ler_variavel(nome: str) -> str:
    valor = os.getenv(nome)
    if not valor:
        raise EnvironmentError(f"Variável de ambiente '{nome}' não definida. Confira o arquivo .env.")
    return valor


PASTA_REDE = Path(ler_variavel("PASTA_REDE"))
NOME_ARQUIVO = ler_variavel("NOME_ARQUIVO")
USUARIO_DB = ler_variavel("USUARIO_DB")
SENHA_DB = ler_variavel("SENHA_DB")

PASTA_RAIZ_DESTINO = PASTA_REDE / "Copias_Teste"

PASTA_BACKUP = PASTA_RAIZ_DESTINO / "01_Backup"
PASTA_ATUALIZADO = PASTA_RAIZ_DESTINO / "02_Atualizado"

ARQUIVO_ORIGINAL = PASTA_REDE / NOME_ARQUIVO
ARQUIVO_SEGURANCA = PASTA_BACKUP / f"{ARQUIVO_ORIGINAL.stem} - Backup{ARQUIVO_ORIGINAL.suffix}"
ARQUIVO_ATUALIZADO = PASTA_ATUALIZADO / f"{ARQUIVO_ORIGINAL.stem} - Atualizado{ARQUIVO_ORIGINAL.suffix}"
ARQUIVO_LOG = PASTA_RAIZ_DESTINO / "log_automacao.txt"


def registrar_log(mensagem: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(ARQUIVO_LOG, "a", encoding="utf-8") as log_file:
            log_file.write(f"[{timestamp}] {mensagem}\n")
    except Exception as e:
        print(f"Falha crítica ao gravar log: {e}")


def iniciar_bloco_log() -> None:
    agora = datetime.now()
    data_str = agora.strftime("%d/%m/%Y")
    hora_str = agora.strftime("%H:%M:%S")

    divisorias = "\n" + "=" * 60 + "\n"
    cabecalho = f"EXECUÇÃO EM: {data_str} ÀS {hora_str}\n"
    linha_final = "-" * 60 + "\n"

    try:
        with open(ARQUIVO_LOG, "a", encoding="utf-8") as log_file:
            log_file.write(divisorias + cabecalho + linha_final)
    except Exception as e:
        print(f"Falha ao iniciar bloco de log: {e}")


def garantir_existencia_da_pasta(caminho_pasta: Path) -> None:
    caminho_pasta.mkdir(parents=True, exist_ok=True)


def duplicar_arquivo(origem: Path, destino: Path, tipo_copia: str) -> None:
    if not origem.exists():
        raise FileNotFoundError(f"Arquivo de origem não encontrado: {origem}")

    shutil.copy2(origem, destino)
    registrar_log(f"Sucesso: Cópia de {tipo_copia} gerada em {destino.parent.name}/{destino.name}")


def ajustar_conexoes_e_credenciais(workbook: xw.Book) -> None:
    try:
        for conn in workbook.api.Connections:
            try:
                if conn.Type == 1:
                    oledb = conn.OLEDBConnection
                    oledb.BackgroundQuery = False
                    conn_str = oledb.Connection
                    if "UID=" not in conn_str and "User Id=" not in conn_str:
                        oledb.Connection = f"{conn_str};User ID={USUARIO_DB};Password={SENHA_DB};"

                elif conn.Type == 2:
                    odbc = conn.ODBCConnection
                    odbc.BackgroundQuery = False
                    conn_str = odbc.Connection
                    if "UID=" not in conn_str and "PWD=" not in conn_str:
                        odbc.Connection = f"{conn_str};UID={USUARIO_DB};PWD={SENHA_DB};"
            except Exception:
                pass
    except Exception:
        pass


def atualizar_dados_excel(caminho_arquivo: Path) -> None:
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.open(caminho_arquivo)

        ajustar_conexoes_e_credenciais(wb)

        inicio_refresh = time.time()
        wb.api.RefreshAll()

        try:
            app.api.CalculateUntilAsyncQueriesDone()
        except Exception:
            time.sleep(15)

        fim_refresh = time.time()
        tempo_decorrido = fim_refresh - inicio_refresh

        if tempo_decorrido >= 60:
            minutos = int(tempo_decorrido // 60)
            segundos = int(tempo_decorrido % 60)
            texto_tempo = f"{minutos} min e {segundos} seg"
        else:
            texto_tempo = f"{tempo_decorrido:.2f} segundos"

        registrar_log(f"Tempo de atualização do banco de dados: {texto_tempo}")

        for nome_guia in ["FRONT", "Analitico"]:
            try:
                sheet = wb.sheets[nome_guia]
                sheet.autofit(axis="columns")
            except Exception as e:
                registrar_log(f"Aviso: Não foi possível auto-ajustar as colunas da guia {nome_guia}: {e}")

        sheet_front = wb.sheets["FRONT"]
        sheet_front.activate()

        wb.save()
        wb.close()
        registrar_log(f"Sucesso: Dados atualizados e colunas formatadas no arquivo {caminho_arquivo.name}")

    except Exception as e:
        registrar_log(f"Erro interno no processamento do Excel: {e}")
        raise e
    finally:
        app.quit()


def executar_automacao_relatorio() -> None:
    try:
        garantir_existencia_da_pasta(PASTA_BACKUP)
        garantir_existencia_da_pasta(PASTA_ATUALIZADO)

        iniciar_bloco_log()
        registrar_log("Início da execução da automação.")

        duplicar_arquivo(ARQUIVO_ORIGINAL, ARQUIVO_SEGURANCA, "Segurança (Backup)")
        duplicar_arquivo(ARQUIVO_ORIGINAL, ARQUIVO_ATUALIZADO, "Produção (Para Atualizar)")

        atualizar_dados_excel(ARQUIVO_ATUALIZADO)

        registrar_log("Fim da execução: Automação concluída com 100% de sucesso.")
        print("Automação concluída com sucesso! Pastas separadas e log estruturado.")

    except FileNotFoundError as e:
        print(f"[ERRO DE CAMINHO] {e}")
        registrar_log(f"Erro de Caminho: {e}")
    except Exception as e:
        print(f"[ERRO INESPERADO] {e}")
        registrar_log(f"Erro Inesperado: {e}")


if __name__ == "__main__":
    executar_automacao_relatorio()
