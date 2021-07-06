import pandas as pd
import gspread
import time
import os
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from google.oauth2.service_account import Credentials

file_ = "/home/leopoldo/Downloads/USD_BRL Dados Históricos.csv"


def input_dados():
    """Função Principal
    Escolhe o tipo de relatório
    Chama a função data
    Chama a função de extrair os dados
    Chama a função de tratar os dados"""

    remove_file(file_)

    print("Digite o número do tipo do relatório desejado:")
    print("1-Diário")
    print("2-Semanal")
    print("3-Mensal")
    resposta = input()

    try:
        # Verificação das respostas
        dicionario_respostas = {'1': 'Daily', '2': 'Weekly', '3': 'Monthly'}
        tipo_relatorio = dicionario_respostas[resposta]
        data_inicial, data_final = pergunta_data()

        # Extração dos dados selenium
        extrair_dados(tipo_relatorio, data_inicial, data_final)

        # Leitura e tratamento dos dados
        tratamento_dados(tipo_relatorio)
        time.sleep(3)

    except KeyError:
        print("Dados inválidos, retornando para o menu inicial")
        input_dados()


def remove_file(file):
    """Remove o arquivo que fará o download"""
    try:
        os.remove(file)
    except OSError as e:
        return(e)
    else:
        return("File is deleted successfully")


def valida_data(data_str):
    """Valida se é uma data válida"""
    try:
        datetime.strptime(data_str, "%d/%m/%Y")
        return data_str
    except:
        print("Digite uma data válida")
        input_dados()


def pergunta_data():
    """Pergunta data inicial e data final do relatório"""
    data_inicial = input("Digite a data inicial")
    data_inicial = valida_data(data_inicial)

    data_final = input("Digite a data final")
    data_final = valida_data(data_final)
    return data_inicial, data_final


def tratamento_dados(tipo_relatorio):
    time.sleep(5)
    # Leitura do arquivo de downwload
    dataset = pd.read_csv("/home/leopoldo/Downloads/USD_BRL Dados Históricos.csv", decimal=",", )


    # Conexão com a google Sheets e tratamento de dados
    worksheet = get_connection(tipo_relatorio,
                               "https://docs.google.com/spreadsheets/d/1-Cgk6oY2eQWxuCWtmNqDWVMI4zy6cCVNPya09oosreA/")
    valores = get_values(worksheet, "A:F")
    datasetfinal = pd.concat([valores, dataset])
    datasetfinal.fillna("", inplace=True)
    datasetfinal["Data"] = datasetfinal["Data"].apply(lambda x: x.replace(".", "/"))
    datasetfinal["Var%"] = datasetfinal["Var%"].apply(lambda x: x.replace("%",""))
    datasetfinal.drop_duplicates("Data", keep="last", inplace=True)

    # Update dos dados
    valores_finais = datasetfinal.values.tolist()
    worksheet.update("A2", valores_finais, value_input_option="USER_ENTERED")
    print("Valores inseridos com sucesso")


def get_connection(name, id_):
    """ Função de conexão com o Gsheets"""
    scope = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets",
             "https://spreadsheets.google.com/feeds"]

    creds = Credentials.from_service_account_file('/home/leopoldo/PycharmProjects/projetos_rpa/chave.json',scopes=scope
    )
    client = gspread.authorize(creds)
    sheet = client.open_by_url(id_)
    worksheet = sheet.worksheet(name)
    return worksheet


def get_values(worksheet, rangetotal):
    """Função para obter os valores da Gsheets"""
    values = worksheet.get(rangetotal)
    values = pd.DataFrame(values)
    values.columns = values.iloc[0]
    values.drop(axis=1, index=0, inplace=True)
    return values


def extrair_dados(tipo_relatorio, data_inicial, data_final):
    """Extração do relatório via selenium"""

    # Configurações
    caps = DesiredCapabilities().CHROME
    caps["pageLoadStrategy"] = "eager"
    driver = webdriver.Chrome(ChromeDriverManager().install(), desired_capabilities=caps)
    driver.implicitly_wait(10)

    # Inicialiozação
    driver.maximize_window()
    base_ini = "https://br.investing.com/currencies/usd-brl-historical-data"
    driver.get(base_ini)
    time.sleep(2)

    # Tirar Popup e login
    popup = driver.find_element_by_xpath("/html//button[@id='onetrust-accept-btn-handler']")
    time.sleep(0.25)
    popup.click()
    time.sleep(0.25)
    driver.find_element_by_xpath("//span[@id='userAccount']/div/a[1]").click()
    time.sleep(0.25)
    driver.find_element_by_xpath("/html//div[@id='signup']/div[@class='signBtnWrap']/div[1]/span[@class='text']").click()

    for h in driver.window_handles:
        driver.switch_to.window(h)

    driver.find_element_by_id("email").send_keys("leozanellato@gmail.com")
    driver.find_element_by_id("pass").send_keys("suasenha")
    driver.find_element_by_name("login").click()

    time.sleep(5)
    for h in driver.window_handles:
        driver.switch_to.window(h)

    time.sleep(5)

    # Início do relatório - diário, semanal, mensal
    select = Select(driver.find_element_by_xpath("/html//select[@id='data_interval']"))
    # select by visible text
    select.select_by_value(tipo_relatorio)

    # Preenchimento calendário
    time.sleep(0.25)
    driver.find_element_by_xpath("//div[@id='flatDatePickerCanvasHol']//span[@class='datePickerIcon']").click()
    time.sleep(0.25)
    driver.find_element_by_xpath("/html//input[@id='startDate']").clear()
    time.sleep(0.25)
    driver.find_element_by_xpath("/html//input[@id='startDate']").send_keys(data_inicial)

    driver.find_element_by_xpath("/html//input[@id='endDate']").clear()
    driver.find_element_by_xpath("/html//input[@id='endDate']").send_keys(data_final)

    time.sleep(0.25)
    driver.find_element_by_xpath("/html//a[@id='applyBtn']").click()
    time.sleep(2)
    driver.find_element_by_xpath("/html//div[@id='column-content']//a[@title='Baixar dados']").click()
    time.sleep(5)
    driver.close()


input_dados()
