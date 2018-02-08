from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, NoAlertPresentException
from selenium import webdriver
import openpyxl
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

#  Acessa os dados de login fora do script, salvo numa planilha existente, para proteger as informações de credenciais
from selenium.webdriver.support.wait import WebDriverWait

dados = openpyxl.load_workbook('C:\\gomnet.xlsx')
login = dados['Plan1']
url = 'http://gomnet.ampla.com/'
urlVistSob = 'http://gomnet.ampla.com/vistoria/vistorias.aspx'
consulta = 'http://gomnet.ampla.com/ConsultaObra.aspx'
username = login['A1'].value
password = login['A2'].value
Revisao = "N"  # Altere para 'S' caso deseje buscar por sobs com status "Pendente", por terem ido para revisão de projeto

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
            buscaVist = driver.window_handles[0]
            # Opção de buscar por sobs com status "Pendente" (foram para revisão de projeto)
            if Revisao == "S":  # TODO Implementar detalhes de vistoria para obras em revisão de projeto
                statusSob = Select(driver.find_element_by_id('ctl00_ContentPlaceHolder1_ddlStatus'))
                statusSob.select_by_visible_text('PENDENTE')
            sob = driver.find_element_by_id('ctl00_ContentPlaceHolder1_txtBoxNumSOB')
            sob.clear()
            sob.send_keys(line)
            try:  # Verifica se a sob foi encontrada, caso contrário, registra no 'log.txt'
                for x in range(0, 3):  # Busca pela sob 03 vezes para ter certeza de que não consta registro
                    try:
                        driver.find_element_by_id('ctl00_ContentPlaceHolder1_ImageButton_Enviar').click()
                        detalheBtn = driver.find_element_by_id(
                            'ctl00_ContentPlaceHolder1_gridViewVistorias_ctl02_ImageButton_DetalhesVistoria')
                        if detalheBtn.is_displayed():
                            break
                    except NoSuchElementException:
                        continue
                driver.find_element_by_id(
                    'ctl00_ContentPlaceHolder1_gridViewVistorias_ctl02_ImageButton_DetalhesVistoria').click()
                detalhesVist = driver.window_handles[1]
                driver.switch_to_window(detalhesVist)
                chkProg = driver.find_element_by_id('chkProgramacao') # Insere checkbox "Programação" numa variável
                if chkProg.is_selected(): # Verifica se a checkbox está selecionada
                    driver.close()
                else:
                    hora = driver.find_element_by_id('txtHora')
                    c = 0
                    while c <= 5:
                        hora.send_keys(Keys.CONTROL + Keys.LEFT)
                        c += 1
                    hora.send_keys('0800')
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'chkProgramacao'))).click()
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'Button2'))).click() # Finaliza a vistoria
                    try:
                        confirmMsg = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                            (By.XPATH, "//*[contains(text(), 'Operação realizada com sucesso')]")))
                        if confirmMsg.is_displayed():
                            print(line + ' vistoriada com sucesso.')
                    except NoSuchElementException:
                        falhaDetalhe = open('log.txt', 'a')
                        falhaDetalhe.write(line + ' falha ao inserir detalhes. Favor verificar')
                        falhaDetalhe.close()
                    driver.close()
                driver.switch_to_window(buscaVist)
            except NoSuchElementException:
                falhaVist = open('log.txt', 'a')
                falhaVist.write(line + ' não encontrada' + '\n')
                falhaVist.close()
                continue
