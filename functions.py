import os
import shutil
from pathlib import Path
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import WebDriverException

# Função para ler o arquivo de parametros


def read_archive(caminho_arquivo: str, chave: str = 'url_site') -> str:
    try:
        with open(caminho_arquivo, 'r') as file:
            for line in file:
                if chave in line:
                    return line.split('=', 1)[1].strip()
    except FileNotFoundError:
        print(f"Arquivo não encontrado: {caminho_arquivo}")
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")

    return None  # Retorna None caso não encontre a chave ou ocorra erro

# Função para validar o ticker


def validar_ticker(ticker):
    ticker = ticker.strip().upper()
    return 5 <= len(ticker) <= 6


# Função para capturar os indicadores
def capturar_dados(driver):
    indicadores = [
        "Cotação", "Min 52 sem", "Max 52 sem", "Valor de mercado",
        "Nro. Ações", "Dia", "P/L", "LPA", "P/VP", "VPA", "ROE"
    ]

    dados = {}
    for indicador in indicadores:
        dados[indicador] = WebIndicadoresJs(driver, indicador)
    return dados


def EscreveLog(mensagem):

    now = datetime.now()

    # Passando fixo para não ter que passar na chamada da função e já valido a criação dessas pastas
    ArquivoLog = f"{Path.home()}\\BPA001 - BuscaCotacoes\\1. LOG\\{now.strftime('%Y')}\\{now.strftime('%m')}\\{now.strftime('%d')}\\LOG.txt"

    dataHora = now.strftime("%Y-%m-%d %H:%M:%S")

    my_file = open(ArquivoLog, 'a', encoding='utf-8')

    my_file.write(f'{dataHora} - ' + f'{mensagem}' + '\n')

    my_file.close()


# Função que faz a validação se as pasta já estão criadas
def ValidaArquivo():

    CaminhoProjeto = f'{Path.home()}\\BPA001 - BuscaCotacoes'

    ArquivoLog = f'{CaminhoProjeto}\\1. LOG'

    CaminhoInput = f'{CaminhoProjeto}\\2. INPUT'

    CaminhoProcessamento = f'{CaminhoProjeto}\\3. PROCESSAMENTO'

    CaminhoFinalizado = f'{CaminhoProjeto}\\4. FINALIZADO'

    # Capturando a data de hoje
    now = datetime.now()

    # Validando se a pasta do caminho do projeto existe
    if os.path.isdir(CaminhoProjeto) == False:

        os.mkdir(CaminhoProjeto)

    # Validando se a pasta de Log existe
    if os.path.isdir(ArquivoLog) == False:

        os.mkdir(ArquivoLog)

    ArquivoLog = f'{ArquivoLog}\\{now.strftime("%Y")}'

    # Criando pasta de acordo com ano, mes e dia
    if os.path.isdir(ArquivoLog) == False:

        os.mkdir(ArquivoLog)

    ArquivoLog = f'{ArquivoLog}\\{now.strftime("%m")}'

    # Criando pasta de acordo com ano, mes e dia
    if os.path.isdir(ArquivoLog) == False:

        os.mkdir(ArquivoLog)

    ArquivoLog = f'{ArquivoLog}\\{now.strftime("%d")}'

    # Criando pasta de acordo com ano, mes e dia
    if os.path.isdir(ArquivoLog) == False:

        os.mkdir(ArquivoLog)

    ArquivoLog = f'{ArquivoLog}\\LOG.txt'

    # Validando se a pasta  Input existe
    if os.path.isdir(CaminhoInput) == False:

        os.mkdir(CaminhoInput)

    # Validando se a pasta Processamento existe
    if os.path.isdir(CaminhoProcessamento) == False:

        os.mkdir(CaminhoProcessamento)

    # Validando se a pasta Finalizado existe
    if os.path.isdir(CaminhoFinalizado) == False:

        os.mkdir(CaminhoFinalizado)

    EscreveLog(
        "=========================== INICIO - Valida Arquivo ================================")

    CaminhoArquivoExcel = ""

    # Validando se já não contem arquivo na pasta PROCESSAMENTO
    mensagem = "Validando se já não contem arquivo na pasta PROCESSAMENTO"
    EscreveLog(mensagem)

    caminhosArquivo = [
        os.path.join(CaminhoProcessamento, nome)
        for nome in os.listdir(CaminhoProcessamento)
    ]

    for arq in caminhosArquivo:
        if arq.lower().endswith(".xlsx"):
            CaminhoArquivoExcel = arq
            mensagem = f"Arquivo encontrado na pasta PROCESSAMENTO: {CaminhoArquivoExcel}"
            EscreveLog(mensagem)

    if len(CaminhoArquivoExcel) == 0:

        # Listando os arquivos dentro da pasta INPUT
        mensagem = "Listando os arquivos dentro da pasta INPUT"
        EscreveLog(mensagem)

        caminhosArquivo = [
            os.path.join(CaminhoInput, nome)
            for nome in os.listdir(CaminhoInput)
        ]

        # Capturando o nome do arquivo excel e movendo para pasta de processamento
        mensagem = "Capturando o nome do arquivo excel e movendo para pasta de processamento"
        EscreveLog(mensagem)

        for arq in caminhosArquivo:
            if arq.lower().endswith(".xlsx"):
                CaminhoArquivoExcel = arq
                shutil.move(CaminhoArquivoExcel, CaminhoProcessamento)

                # Movendo arquivo de INPUT para PROCESSAMENTO
                mensagem = f"Movendo arquivo de INPUT: {CaminhoArquivoExcel} para PROCESSAMENTO: {CaminhoProcessamento}"
                EscreveLog(mensagem)

                CaminhoArquivoExcel = arq.replace(
                    "2. INPUT", "3. PROCESSAMENTO")
                break

        EscreveLog(
            "=========================== FIM - Valida Arquivo ================================")

    return CaminhoArquivoExcel


# Função para abrir o navegador


def AbreNavegador(Url):
    options = webdriver.ChromeOptions()
    options.add_argument('--log-level=3')
    options.add_argument("--headless=new")  # Modo headless
    driver = webdriver.Chrome(service=Service(
        ChromeDriverManager().install()), options=options)
    driver.get(Url)
    return driver

# Navegar por URL


def webNavegarUrl(url, driver):
    driver.get(url)


# Fechando a sessão do driver


def FecharNavegador(driver):
    driver.close()

# Função para executar comando JS


def ExecutaJs(script, driver):
    driver.execute_script(script)

# Função para executar um JS e ter o retorno do valor


def WebRetornaJs(script, driver):
    WebRetornaJs = driver.execute_script(script)
    return WebRetornaJs

# Fução para clicar pelo Id


def ClickId(Id, driver):

    driver.execute_script("document.getElementById('"+Id+"').click()")

# Função para setar o elemento na pagina


def SetaElementoId(Id, Valor, driver):

    driver.execute_script(
        "document.getElementById('"+Id+"').value='"+Valor+"'")

# Função para capturar o texto por Js e fazer a validação de carragemento da tela


def WebValidaTextJs(Id, tempo, TextoElemento, driver, time, TextoErro="Não há dados a serem exibidos"):

    i = 0

    for i in range(tempo):

        # Capturando o valor do texto
        try:
            ValidaCarragamento = driver.execute_script(
                "return document.getElementById('"+Id+"').innerText")
        except:
            pass

        # Validando se o elemento foi encontrado
        if ValidaCarragamento == TextoElemento:
            break

            # Validando se o elemento foi encontrado
        if ValidaCarragamento == TextoErro:
            break

        # Validando se já deu o tempo de validação
        if i >= tempo:

            break

        time.sleep(1)

    return ValidaCarragamento

# Função para capturar o valor do texto por Js


def WebGetTextJs(Id, driver):

    WebGetTextJs = driver.execute_script(
        "return document.getElementById('"+Id+"').innerText")

    return WebGetTextJs

# Função para capturar indicadores


def WebIndicadoresJs(driver, indicador):
    script = f"""
    function getValueByLabel(label) {{
        var tds = document.querySelectorAll('td');
        for (var i = 0; i < tds.length; i++) {{
            if (tds[i].innerText.trim() === label) {{
                return tds[i + 1] ? tds[i + 1].innerText.trim() : null;
            }}
        }}
        return null;
    }}
    return getValueByLabel("{indicador}");
    """
    return driver.execute_script(script)
