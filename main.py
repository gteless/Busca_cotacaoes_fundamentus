# import das libs
import pyodbc
import shutil
from datetime import datetime

# import das functions
from functions import *

# variaveis
parametros = 'parameters.txt'
CaminhoArquivoExcel = ValidaArquivo()
CaminhoProjeto = f'{Path.home()}\\BPA001 - BuscaCotacoes'
CaminhoFinalizado = f'{CaminhoProjeto}\\4. FINALIZADO'
url_pesquisa = read_archive(parametros, 'url_pesquisa')
homepage_url = read_archive(parametros, 'url_home')


if len(CaminhoArquivoExcel) > 0:

    # Capturando driver
    mensagem = "Capturando driver"
    EscreveLog(mensagem)

    for driver in pyodbc.drivers():

        # Pegando o nome apenas para o driver .xlsx
        mensagem = "Pegando o nome apenas para o driver .xlsx"
        EscreveLog(mensagem)

        if '.xlsx' in driver:
            myDriver = driver

    # Definindo connection string
    mensagem = "Definindo connection string"
    EscreveLog(mensagem)

    conn_str = (r'DRIVER={' + myDriver + '};'
                f'DBQ={CaminhoArquivoExcel};'
                r'ReadOnly=0')  # O padrão do Excel é uma conexão somente leitura, portanto, se você quiser atualizar a planilha, inclua ReadOnly=0

    # Definir nossa conexão, autocommit DEVE SER CONFIGURADO PARA TRUE, também podemos editar dados.
    cnxn = pyodbc.connect(conn_str, autocommit=True)
    cursor_select = cnxn.cursor()
    cursor_update = cnxn.cursor()

    for worksheet in cursor_select.tables():

        # Pegando worksheet
        mensagem = "Pegando worksheet"
        EscreveLog(mensagem)
        tableName = worksheet[2]

    # "SELECT * FROM [Planilha1$]"
    mensagem = f"Query executada: SELECT * FROM [{format(tableName)}] WHERE [Status] IS NULL"
    EscreveLog(mensagem)

    Query = f"SELECT * FROM [{format(tableName)}] WHERE [Status] IS NULL"
    cursor_select.execute(Query)

    if homepage_url:
        # Abrindo navegador
        mensagem = f"Abrindo navegador na página - ({homepage_url})"
        EscreveLog(mensagem)
        try:
            driver = AbreNavegador(homepage_url)

        except WebDriverException as e:
            # Ocorreu um erro ao tentar acessar a URL
            mensagem = f"Ocorreu um erro ao tentar acessar a URL: {e}"
            EscreveLog(mensagem)
            exit(1)
    else:
        # Erro: URL base não encontrada no arquivo de parâmetros.
        mensagem = "Erro: URL base não encontrada no arquivo de parâmetros."
        EscreveLog(mensagem)
        exit(1)

    # Loop na minha tabela
    for row in cursor_select:

        # Setando variaveis
        ticker = row.Ticker
        if url_pesquisa:
            url = url_pesquisa.replace('[]', ticker)
        else:
            # Erro: URL base não encontrada no arquivo de parâmetros.
            mensagem = "Erro: URL base não encontrada no arquivo de parâmetros."
            EscreveLog(mensagem)
            exit(1)

        # Consultado o Ticker {Ticker} no site: {url}
        mensagem = f"Consultado o Ticker {ticker} no site: {url}"
        EscreveLog(mensagem)
        webNavegarUrl(url, driver)

        # Verifica se carregou a página com o Ticker informado pelo usuário
        mensagem = "Verifica se carregou a página com o Ticker informado pelo usuário"
        EscreveLog(mensagem)
        script = "return document.getElementsByTagName('td')[1].innerText"
        retorno = WebRetornaJs(script, driver)
        if retorno.upper() == ticker.upper():
            # Página carregada com sucesso, inciando a captura dos dados
            mensagem = "Página carregada com sucesso, iniciando a captura dos dados"
            EscreveLog(mensagem)
        else:
            # Erro: Página não carregou com o Ticker esperado
            mensagem = "Erro: Página não carregou com o Ticker esperado"
            EscreveLog(mensagem)
            exit(1)

        # capturando os dados
        dados = capturar_dados(driver)
        now = datetime.now()

        # Dados capturados
        mensagem = f"Dados capturados: {' | '.join([f'{key}: {value}' for key, value in dados.items()])}"
        EscreveLog(mensagem)

       # Devido no site o nome desses campos terem caracteres especiais, aqui estou tratando
       # para na hora de fazer o update na planilha não ter erros
        colunas_renomeadas = {
            "Cotação": "Cotacao",
            "Nro. Ações": "Nro Acoes"
        }

        # Crio um novo dicionário corrigindo os valores, caso tenham outros, é possível adicionar
        dados_tratados = {}
        for chave, valor in dados.items():
            # Se não estiver no dicionário, mantém o original
            nova_chave = colunas_renomeadas.get(chave, chave)
            dados_tratados[nova_chave] = valor

        # Para depois trabalhar no mesmo dicionário, estou adicionando as colunas "extras"
        dados_tratados["Ticker"] = ticker
        dados_tratados["Status"] = "OK"
        dados_tratados["Data da Consulta"] = now.strftime("%Y-%m-%d %H:%M:%S")

        # Atribui as variáveis a partir dos dados tratados
        cotacao = dados_tratados.get("Cotacao", None)
        min_52_sem = dados_tratados.get("Min 52 sem", None)
        max_52_sem = dados_tratados.get("Max 52 sem", None)
        valor_mercado = dados_tratados.get("Valor de mercado", None)
        nro_acoes = dados_tratados.get("Nro Acoes", None)
        dia = dados_tratados.get("Dia", None)
        pl = dados_tratados.get("P/L", None)
        lpa = dados_tratados.get("LPA", None)
        pvp = dados_tratados.get("P/VP", None)
        vpa = dados_tratados.get("VPA", None)
        roe = dados_tratados.get("ROE", None)
        data_consulta = dados_tratados.get("Data da Consulta", None)
        status = dados_tratados.get("Status", None)

        Query = f"""
            UPDATE [{format(tableName)}] 
            SET [Cotacao] = '{cotacao}', 
                [Min 52 sem] = '{min_52_sem}', 
                [Max 52 sem] = '{max_52_sem}', 
                [Valor de mercado] = '{valor_mercado}', 
                [Nro Acoes] = '{nro_acoes}', 
                [Dia] = '{dia}',
                [P/L] = '{pl}', 
                [LPA] = '{lpa}', 
                [P/VP] = '{pvp}', 
                [VPA] = '{vpa}', 
                [ROE] = '{roe}',
                [Data da Consulta] = '{data_consulta}',
                [Status] = '{status}' 
            WHERE [Ticker] = '{ticker}'
        """

        # Executa a query com os parâmetros
        cursor_update.execute(Query)

    cursor_update.close()
    cursor_select.close()
    cnxn.close()

    # Matando sessão do driver no fim do processo
    mensagem = "Matando sessão do driver no fim do processo"
    EscreveLog(mensagem)
    FecharNavegador(driver)

    # Movendo arquivo para de finalizado
    mensagem = "Movendo arquivo para de finalizado"
    EscreveLog(mensagem)
    shutil.move(CaminhoArquivoExcel, CaminhoFinalizado)

    EscreveLog(
        "=========================== FIM - Navegação Busca Cotações ================================")
