# Automação de Relatórios Excel com SQL Server

Script em Python que atualiza automaticamente relatórios em Excel conectados a um banco SQL Server, com backup do arquivo original e log de cada execução.

Criei a primeira versão no meu trabalho como Jovem Aprendiz para eliminar a atualização manual de relatórios. Esta é uma versão sem dados da empresa: caminhos, nomes de arquivos e credenciais ficam em variáveis de ambiente.

## O que o script faz

1. Cria as pastas de backup e de arquivos atualizados, se ainda não existirem.
2. Faz uma cópia de segurança do relatório original antes de qualquer alteração.
3. Abre uma segunda cópia no Excel em segundo plano, com `xlwings`.
4. Injeta as credenciais do banco nas conexões OLEDB/ODBC da planilha.
5. Atualiza todas as consultas ao SQL Server (`RefreshAll`) e mede o tempo gasto.
6. Ajusta a largura das colunas das abas de saída e salva o arquivo.
7. Registra tudo em um arquivo de log com data, hora e resultado.

## Tecnologias

- Python 3
- xlwings (controle do Microsoft Excel)
- SQL Server (via conexões OLEDB/ODBC do Excel)
- python-dotenv (leitura das variáveis de ambiente)

## Como rodar

Requisitos: Windows com Microsoft Excel instalado e uma planilha com conexões ao SQL Server.

```bash
pip install xlwings python-dotenv
```

Copie o arquivo `.env.example` para `.env` e preencha:

```
PASTA_REDE=C:\caminho\da\pasta
NOME_ARQUIVO=Relatorio.xlsb
USUARIO_DB=seu_usuario
SENHA_DB=sua_senha
```

Depois execute:

```bash
python automacao_relatorio.py
```

O arquivo `.env` está no `.gitignore` e nunca deve ser enviado ao repositório.

## Próximas melhorias

- Registrar no log os erros de conexão que hoje são ignorados.
- Trocar a espera fixa de 15 segundos por uma verificação real do fim das consultas.
- Enviar o relatório atualizado por e-mail automaticamente.

## Autor

Enzo Guedes Moura · [GitHub](https://github.com/Kfhw460) · enzoguedes460@gmail.com

## Inteligência Artificial para uso de aprendizagem e correção.
