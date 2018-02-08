from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, NoAlertPresentException
from selenium import webdriver
import openpyxl
from selenium.webdriver.support.ui import Select
import time

#  Acessa os dados de login fora do script, salvo numa planilha existente, para proteger as informações de credenciais
dados = openpyxl.load_workbook('C:\\gomnet.xlsx')
login = dados['Plan1']
url = 'http://gomnet.ampla.com/'
urlVistSob = 'http://gomnet.ampla.com/vistoria/vistorias.aspx'
consulta = 'http://gomnet.ampla.com/ConsultaObra.aspx'
username = login['A1'].value
password = login['A2'].value
Revisao = "N" # Altere para 'S' caso deseje buscar por sobs com status "Pendente", por terem ido para revisão de projeto

# --------------- Headless Mode -------------------------
# chromeOptions = webdriver.ChromeOptions()
# prefs = {"download.default_directory" : os.getcwd(),
#          "download.prompt_for_download": False}
# chromeOptions.add_experimental_option("prefs",prefs)
# chromeOptions.add_argument('--headless')
# chromeOptions.add_argument('--window-size= 1600x900')
# driver = webdriver.Chrome(chrome_options=chromeOptions)
# -------------------------------------------------------

driver = webdriver.Chrome()

if __name__ == '__main__':
    driver.get(url)
    # Insere login/senha e entra no sistema
    uname = driver.find_element_by_name('txtBoxLogin')
    uname.send_keys(username)
    passw = driver.find_element_by_name('txtBoxSenha')
    passw.send_keys(password)
    submit_button = driver.find_element_by_id('ImageButton_Login').click()

    # Acessa a página de Vistoria de Obra
    driver.get(urlVistSob)

    # Insere o número da Sob em seu respectivo campo e realiza a busca
    with open('sobs.txt') as data:
        datalines = (line.strip('\r\n') for line in data)
        for line in datalines:
            # Opção de buscar por sobs com status "Pendente" (foram para revisão de projeto)
            if Revisao == "S":
                statusSob = Select(driver.find_element_by_id('ctl00_ContentPlaceHolder1_ddlStatus'))
                statusSob.select_by_visible_text('PENDENTE')
            driver.find_element_by_id('ctl00_ContentPlaceHolder1_txtBoxNumSOB').clear()
            sob = driver.find_element_by_id('ctl00_ContentPlaceHolder1_txtBoxNumSOB')
            sob.send_keys(line)
            driver.find_element_by_id('ctl00_ContentPlaceHolder1_ImageButton_Enviar').click()