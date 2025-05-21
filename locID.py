from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
import pandas as pd
import time


# planilha Excel com nome e id
name_planilha = input('Insira aqui o nome da planilha:  ').lower()
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
if 'ID' not in planilha.columns:
    planilha['ID'] = ""

planilha['ID'] = planilha['ID'].astype(str)

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
    find_name = linha['Nome']
    find_cpf = linha['CPF']

    # URL driver list
    dynamic_url = f"{base_url}logistic-operator/driver-list/"
    driver.get(dynamic_url)

    # Buscar id pelo CPF
    try:
            caixa_pesquisa = driver.find_element(By.XPATH, '//*[@id="cpf"]')
            cpf = find_cpf
            caixa_pesquisa.send_keys(cpf)
            driver.find_element(By.XPATH, '//*[@id="root"]/div[2]/main/div/div[1]/div[1]/div/button').click()

            # tempo para busca de cada driver
            WebDriverWait(driver, 10).until(
                 EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div[2]/main/div/div[3]/table/tbody/tr/td[1]'))
            )

            id_driver = driver.find_element(By.XPATH, '//*[@id="root"]/div[2]/main/div/div[3]/table/tbody/tr/td[1]').text

            planilha.at[indice, 'ID'] = id_driver

            print(f'{indice} - Nome: {find_name} - ID encontrado')

    except TimeoutException as e:
        print(f'{indice} - {find_name} - ID não encontrado... ')
        planilha.at[indice, 'ID'] = 'Não está na Franquia'
    except NoSuchElementException as e:
        print(f'{indice} - {find_name} - ID não encontrado... ')
        planilha.at[indice, 'ID'] = 'Não está na Franquia'
         
planilha.to_excel('Driver_ID.xlsx', index= False)
print()
input('Aperte ENTER para fechar o navegador... ')
driver.quit()