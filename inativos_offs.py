from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, StaleElementReferenceException
from selenium.webdriver.chrome.service import Service
import pandas as pd
import time
import sys
import os

# Caminho relativo ao executável
chromedriver_path = os.path.join(os.path.dirname(__file__), 'chromedriver.exe')
service = Service(executable_path=chromedriver_path)
driver = webdriver.Chrome(service=service)

# Abrindo navegador e entregô
complemento = "login/password"
base_url = f"https://franqueado.entregolog.com/"
base_url2 = base_url + complemento
driver.get(base_url2)

# Verifica se o script está rodando como executável
is_executable = hasattr(sys, 'frozen')

def barra_progresso_login():
    try:
        WebDriverWait(driver, 600).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div[2]/main/div/div[1]/h1'))
        )
        print()
        print('Vamos iniciar a raspagem de dados')
    except NoSuchElementException as e:
        print(f'Não foi possível fazer o login: {e}')

# Tempo para login manual (ajuste conforme necessário)
barra_progresso_login()

print()

# Criando um DataFrame vazio para acumular os dados
colunas_inativos = ['Nome', 'cpf', 'subpraca', 'modal', 'maquininha', 'conexao', 'estado de trabalho', 'inativacao']
colunas_offlines = ['Nome', 'cpf', 'subpraca', 'modal', 'maquininha', 'situacao', 'inativacao']

# Criação de DataFrame 
planilha_final = pd.DataFrame(columns=colunas_inativos)

# Função para coleta de dados de inativos
def info_inativos():
    dados = []
    try:
        for indice in range(1, 11):
            nome = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[1]/div/div[1]/div/div[1]/div/span').text
            cpf = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[1]/div/div[2]/div/div[1]/div/span').text
            subpraca = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[2]').text
            modal = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[3]').text
            maquininha = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[4]').text
            conexao = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[5]').text
            estado_de_trabalho = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[6]').text
            inativacao = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[7]').text

            dados.append([nome, cpf, subpraca, modal, maquininha, conexao, estado_de_trabalho, inativacao])
            print(f'{indice} | {nome} | {inativacao}')
    except NoSuchElementException as e:
        print(f"Erro ao buscar informações de inativos: {e}")
    
    return dados

# Função para coleta de dados de offlines
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
            inativacao = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[1]/table/tbody/tr[{indice}]/td[6]').text

            # Preenchendo a coluna 'inativacao' com valor nulo
            dados.append([nome, cpf, subpraca, modal, maquininha, situacao, inativacao, None])
            print(f'{indice} | {nome} | {situacao}')
    except NoSuchElementException as e:
        print(f"Erro ao buscar informações de offlines: {e}")
    
    return dados

try:
    # entrar na pagina inativos
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

    def main():
        global planilha_final  # Declarando planilha_final como global
        for t in range(3, 11):  # Looping para coletar dados de inativos
            dados = info_inativos()
            planilha_final = pd.concat([planilha_final, pd.DataFrame(dados, columns=colunas_inativos)], ignore_index=True)
            time.sleep(8)  # Pausa antes de pular para a próxima página
            
            try:
                next_button = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[2]/nav/ul/li[{t}]/button')
                button_class = next_button.get_attribute('class')
                if 'disabled' in button_class:
                    print("O botão 'Next page' está desativado. Finalizando a coleta de dados de inativos.")
                    break
                else:
                    next_button.click()
                    print()
                    print(f"Página {t-1} de inativos processada.")
                    print()
                    time.sleep(10)
            except NoSuchElementException as e:
                print("Finalizou as páginas de inativos.")
                break

        time.sleep(10)

        # Pular para página dos offlines
        try:
            offs = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div[2]/main/div/div[2]/div[1]/div/div/div[2]/button'))
            )
            actions = ActionChains(driver)
            actions.move_to_element(offs).perform()
            offs.click()
        except TimeoutException as e:
            print(f'O botão "Offlines" não ficou disponível a tempo: {e}')
        except ElementClickInterceptedException as e:
            print(f'O clique foi interceptado por outro elemento: {e}')

        time.sleep(5)

        for i in range(3, 21):
            dados = info_offlines()
            planilha_final = pd.concat([planilha_final, pd.DataFrame(dados, columns=colunas_inativos)], ignore_index=True)
            time.sleep(8)
            
            try:
                next_button = driver.find_element(By.XPATH, f'//*[@id="root"]/div[2]/main/div/div[2]/div[2]/div[2]/nav/ul/li[{i}]/button')
                button_class = next_button.get_attribute('class')
                if 'disabled' in button_class:
                    print("O botão 'Next page' está desativado. Finalizando a coleta de dados de offlines.")
                    break
                else:
                    next_button.click()
                    print()
                    print(f"Página {i-1} de offlines processada.")
                    print()
                    time.sleep(10)
            except NoSuchElementException as e:
                print()
                print("Finalizou as páginas de offlines. ")
                break
    main()

finally:
    # Verifica se planilha_final foi definida antes de tentar salvar
    if not planilha_final.empty:
        planilha_final.to_excel('Planilha_final.xlsx', index=False)
        print()
        print('Planilha Salva')
        print()
    else:
        print("Nenhum dado foi coletado.")

restart = input('Quer tentar novamente? (sim/nao) ')
if restart == 'sim':
    page_inativos = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div[2]/main/div/div[2]/div[1]/div/div/div[1]/button'))
    )
    actions = ActionChains(driver)
    actions.move_to_element(page_inativos).perform()
    page_inativos.click()
    time.sleep(5)
    main()
else:
    input('Aperte ENTER e feche o programa...')

if __name__ == '__main__':
    main()

