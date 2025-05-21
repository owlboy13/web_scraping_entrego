from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
import pandas as pd
import time

# planilha Excel com nome e id
name_planilha = input('Insira aqui o nome da planilha:  ')
format = '.xlsx'
planilha_format = name_planilha + format
planilha = pd.read_excel(planilha_format)

print()

# Inicializando o driver do Chrome
driver = webdriver.Chrome()
complemento = "login/password"
base_url = f"https://franqueado.entregolog.com/"
base_url2 = base_url + complemento
driver.get(base_url2)

# Cria a coluna se não existir
if 'Nome' not in planilha.columns:
    planilha['Nome'] = ""
if 'Modal' not in planilha.columns:
    planilha['Modal'] = ""
if 'Data de nascimento' not in planilha.columns:
    planilha['Data de nascimento'] = ""
if 'Telefone' not in planilha.columns:
    planilha['Telefone'] = ""
if 'CPF' not in planilha.columns:
    planilha['CPF'] = ""
if 'email' not in planilha.columns:
    planilha['email'] = ""
if 'rg' not in planilha.columns:
    planilha['rg'] = ""
if 'cnh' not in planilha.columns:
    planilha['cnh'] = ""

planilha['Nome'] = planilha['Nome'].astype(str)
planilha['Modal'] = planilha['Modal'].astype(str)
planilha['Data de nascimento'] = planilha['Data de nascimento'].astype(str)
planilha['Telefone'] = planilha['Telefone'].astype(str)
planilha['CPF'] = planilha['CPF'].astype(str)
planilha['email'] = planilha['email'].astype(str)
planilha['rg'] = planilha['rg'].astype(str)
planilha['cnh'] = planilha['cnh'].astype(str)

 # aguardar pagina inicial entrego
def carregamento():
    try:
        WebDriverWait(driver, 600).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div[1]/div[1]/button')))
    except NoSuchElementException as e:
        print(f'Não Carregou a página: {e}')

carregamento()

# Iterando pelas linhas da planilha
for indice, linha in planilha.iterrows():
    # find_name = linha['NOME']
    find_id = linha['ID']
    print(f'{indice} | ID:{find_id}')

    # URL driver list
    dynamic_url = f"{base_url}logistic-operator/driver-list/{find_id}"
    driver.get(dynamic_url)

    #Buscar dados dos drivers
    try:
        nome_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Nome completo"]')))
        nome_value = nome_element.get_attribute("value")
        planilha.at[indice, 'Nome'] = nome_value

        modal_element = driver.find_element(By.XPATH, '//*[@id="Modal atual"]')
        modal_value = modal_element.get_attribute("value")
        planilha.at[indice, 'Modal'] = modal_value

        data_nascimento = driver.find_element(By.XPATH, '//*[@id="Data de nascimento"]')
        data_value = data_nascimento.get_attribute("value")
        planilha.at[indice, 'Data de nascimento'] = data_value

        telefone = driver.find_element(By.XPATH, '//*[@id="Telefone (com DDD)"]')
        telefone_value = telefone.get_attribute("value")
        planilha.at[indice, 'Telefone'] = telefone_value

        cpf = driver.find_element(By.XPATH, '//*[@id="CPF"]')
        cpf_value = cpf.get_attribute("value")
        planilha.at[indice, 'CPF'] = cpf_value

        email = driver.find_element(By.XPATH, '//*[@id="E-mail"]')
        email_value = email.get_attribute("value")
        planilha.at[indice, 'email'] = email_value

        rg = driver.find_element(By.XPATH, '//*[@id="RG"]')
        rg_value = rg.get_attribute("value")
        planilha.at[indice, 'rg'] = rg_value

        cnh = driver.find_element(By.XPATH, '//*[@id="CNH"]')
        cnh_value = cnh.get_attribute("value")
        planilha.at[indice, 'cnh'] = cnh_value

    except (NoSuchElementException, TimeoutException) as e:
        print(f"Erro ao capturar informações para o ID {find_id}: {e}")
        planilha.at[indice, 'Nome'] = "não encontrado"
        planilha.at[indice, 'Modal'] = "não encontrado"
        planilha.at[indice, 'Data de nascimento'] = "não encontrado"
        planilha.at[indice, 'Telefone'] = "não encontrado"
        planilha.at[indice, 'CPF'] = "não encontrado"
        planilha.at[indice, 'email'] = "não encontrado"
        planilha.at[indice, 'rg'] = "não encontrado"
        planilha.at[indice, 'cnh'] = "não encontrado"

    # Tempo para garantir que a página carregue completamente
    time.sleep(2)

# Salvando a planilha atualizada
planilha.to_excel('dados_atualizados.xlsx', index=False)

# Aguarde o input do usuário para fechar o script (remova se não necessário)
input('Aperte ENTER para sair...')
driver.quit()
