import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import chromedriver_autoinstaller
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep


# Entrar na planilha e extrair dados


lista_chamados = openpyxl.load_workbook('tickets.xlsx')
pagina_lista_chamados = lista_chamados['Plan1']

driver = webdriver.Chrome()
driver.get('https://csm2.serviceaide.com/#login')
sleep(2)
def logar():

    login = driver.find_element(By.XPATH, "//input[@placeholder='USER NAME']")
    login.send_keys("Gabriel.siqueira@dasa.com.br")
    login.send_keys(Keys.TAB)

    senha = driver.find_element(By.XPATH, "//input[@name='password']")
    senha.send_keys("Ezio@uditore20")

    confirmar = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'SIGN IN')]"))
    )
    confirmar.click()
    sleep(0.5)

def testar_status():
    status_chamado = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[2]/div[1]/div/div[2]/div[3]/div/div[2]/div/div/div/div[2]/div[2]/div[1]/div[2]/div[2]/div[6]/div/div/div[3]"))
    )
    return status_chamado

def pesquisar_chamado():
    pesquisa = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Pesquisar')]"))
    )
    pesquisa.click()
    sleep(0.5)

    consulta = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//input[@autoidentifier='txtbox_GSSearchServiceDesk']"))
    )
    consulta.clear()
    consulta.send_keys(chamado)
    consulta.send_keys(Keys.ENTER)
        
    caso = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//a[@tabindex='1']"))
    )
    caso.click()

def atender_chamado():
    sleep(0.5)
    apontamento = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Apontamentos') and contains(@class, 'x-tab-inner')]"))
    )
    apontamento.click()
    sleep(0.5)

    combo_box_apontamento = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//input[@autoidentifier='combo_WorklogType']"))
    )
    combo_box_apontamento.click()
    sleep(0.5)

    tipo_apontamento = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//li[contains(text(),'Atualizar status')]"))
    )
    tipo_apontamento.click()
    sleep(0.5)

    texto_apontamento = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//textarea[@autoidentifier=('txtbox_Worklog')]"))
    )
    texto_apontamento.send_keys("Chamado encerrado conforme inventário 2023.")
    sleep(0.5)

    visivel_cliente = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//input[@autoidentifier=('chkbox_ClientViewable')]"))
    )
    visivel_cliente.click()
    sleep(0.5)
    acao.click()
    sleep(0.5)

    atualizar_cadastro = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Atualizar cadastro')]"))
    )
    atualizar_cadastro.click()
    sleep(1)
    resolver_chamado()

def resolver_chamado():

    acao.click()
    sleep(0.5)
    retomar = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Retomar solicitação')]"))
    )
    
    retomar.click()
    sleep(1)

    resolucao = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//textarea[@autoidentifier='txtbox_Resolution']"))
    )

    resolucao.clear()
    sleep(1)
    resolucao.send_keys("Chamado encerrado conforme inventário 2023.")
    sleep(1)
    resolucao.send_keys(Keys.TAB)
    acao.click()
    sleep(0.5)

    resolver = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Resolver solicitação')]"))
    )
    
    resolver.click()
    sleep(2)

    subir_aba = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//button[@id='back-btnEl']"))
    )
    subir_aba.click()
    resolucao.text
    alimentar_nova_planilha()
    fechar_chamado()
    
def fechar_chamado():
    sleep(0.5)
    fechar_aba = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//a[@class='x-tab-close-btn']"))
    )
    fechar_aba.click()  

def alimentar_nova_planilha():
    resolucao = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//textarea[@autoidentifier='txtbox_Resolution']"))
    )
    conferencia = resolucao.text
    tickets_concluidos = openpyxl.load_workbook('tickets_concluidos.xlsx')
    pagina_tickets = tickets_concluidos['Plan1']
    pagina_tickets.append([chamado,"Concluido", conferencia , status_chamado.text ,motivo, descrição, Abertoem, Atualizado, Prioridade, Solicitante, Organizaçãodosolicitante, Atribuído, Grupoatribuído, Tipodeticket, Resolverem, StatusdoSLA, Organizaçãodobeneficiário, SelecioneaUnidade, Categorização] )
    tickets_concluidos.save('tickets_concluidos.xlsx')


logar()

for linha in pagina_lista_chamados.iter_rows(min_row=2, values_only=True):
    chamado, status, motivo, descrição, Abertoem, Atualizado, Prioridade, Solicitante, Organizaçãodosolicitante, Atribuído, Grupoatribuído, Tipodeticket, Resolverem, StatusdoSLA, Organizaçãodobeneficiário, SelecioneaUnidade, Categorização = linha

    pesquisar_chamado()
    status_chamado = testar_status()

    if status_chamado.text == "Resolvido":
        alimentar_nova_planilha()
        fechar_chamado()
        continue

    acao = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//button[@id='ca-global-actions-menu-btnEl']"))
    )
        
    if status_chamado.text == "Pendente Fornecedor":
        
        acao.click()
        sleep(0.5)
        retomar = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Retomar solicitação')]"))
    )
        retomar.click()
        atender_chamado()
    
    elif status_chamado.text == "Cadastro contábil":
        resolver_chamado()

    elif status_chamado.text == "Aceitação":

        acao.click()
        sleep(0.5)
        aceitar_solicitacao = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Aceitar solicitação')]"))
        )
        atender_chamado()

    elif status_chamado.text == "Pendente Material":

        acao.click()
        sleep(0.5)
        retomar = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Retomar solicitação')]"))
        )
        retomar.click()
        atender_chamado()

    elif status_chamado.text == "Resposta do Cliente":

        acao.click()
        sleep(0.5)
        aceitar_solicitacao = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Aceitar solicitação')]"))
        )
        atender_chamado()
    
    elif status_chamado.text == "Execução pendente":
        acao.click()
        sleep(0.5)

        retomar = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Retomar solicitação')]"))
        )
        retomar.click()
        atender_chamado()

    elif status_chamado.text == "Em andamento":
        atender_chamado()
