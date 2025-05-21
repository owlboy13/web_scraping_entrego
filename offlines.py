from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
import pandas as pd
import time

# Criando um DataFrame vazio para acumular os dados
planilha = pd.DataFrame(columns=['Nome', 'cpf', 'subpraca', 'modal', 'maquininha','situacao','estado de trabalho'])

# Inicializando o driver do Chrome
driver = webdriver.Chrome()
complemento = "login/password"
base_url = f"https://franqueado.entregolog.com/"
base_url2 = base_url + complemento
driver.get(base_url2)

time.sleep(90)  # Tempo para login manual (ou ajuste conforme necessário)

def info_offlines():
    dados = []
    try:
        for indice in range(1, 11):
            nome = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[1]/div/div[1]/div/div[1]/div/span').text
            cpf = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[1]/div/div[2]/div/div[1]/div/span').text
            subpraca = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[2]').text
            modal = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[3]').text
            maquininha = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[4]').text
            situacao = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[5]').text
            estado_de_trabalho = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[6]').text

            # Adiciona os dados em uma lista
            dados.append([nome, cpf, subpraca, modal, maquininha, situacao, estado_de_trabalho])

            print(f'{indice} | {nome} | {estado_de_trabalho}')
            
    except NoSuchElementException as e:
        print(f"Erro ao buscar informações: {e}")
    
    # Retorna os dados coletados
    return dados

try:
    # Entrando na página de status
    status_frota = f"{base_url}logistic-operator/logistic-operators-driver-status"
    driver.get(status_frota)

    # Filtro e Cidade
    driver.find_element(By.XPATH, '//*[@id="root"]/div[2]/main/div/div[1]/div/button').click()
    caixa_mensagem = driver.find_element(By.XPATH, '//*[@id="downshift-0-input"]')
    cidade = "Sao Paulo"

    time.sleep(3)

    caixa_mensagem.send_keys(cidade)
    for i in range(2):
        caixa_mensagem.send_keys(Keys.ENTER)

    time.sleep(7)

    # Aguardar até o botão "Offlines" estar presente e clicável
    try:
        offs = driver.find_element(By.XPATH, '//*[@id="root"]/div[2]/main/div/div[2]/div[1]/div/div/div[2]/button')
        offs.click()
    except TimeoutException as e:
        print(f'O botão "Offlines" não ficou disponível a tempo: {e}')
    except NoSuchElementException as e:
        print(f'Não encontrou o elemento "Offlines": {e}')

    time.sleep(5)

    # Loop para navegar pelas páginas e coletar os dados
    for i in range(3, 20):  # Ajuste o intervalo conforme necessário
        dados = info_offlines()  # Coletando os dados da página atual
        # Adiciona os dados coletados à planilha
        planilha = pd.concat([planilha, pd.DataFrame(dados, columns=planilha.columns)], ignore_index=True)
        time.sleep(8)
        
        try:
            # Pula para a próxima página
            jump_page = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[2]/nav/ul/li[{i}]/button/div'))
            )
            jump_page.click()
            time.sleep(15)
        except TimeoutException as e:
            print(f"Status Page Finalizado: {e}")
            break

except TimeoutException as e:
    print(f'Tempo de carregamento de página esgotado: {e}')

# Salvando a planilha no Excel
planilha.to_excel("Offlines.xlsx", index=False)
print()
print('Planilha Criada')
print()
input('Aperte ENTER para fechar o programa...')
