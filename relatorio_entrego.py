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
import datetime as dt
import zipfile as zp
import openpyxl
import os

# Formatando data de ontem
today = dt.datetime.now()
yesterday = today - dt.timedelta(days=1)
last_day = yesterday.day

# Inicializando o driver do Chrome
driver = webdriver.Chrome()
complemento = "login/password"
base_url = f"https://franqueado.entregolog.com/"
base_url2 = base_url + complemento
driver.get(base_url2)

# Aguardar pagina inicial entrego
def carregamento():
    try:
        WebDriverWait(driver, 600).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div[1]/div[1]/button')))
    except NoSuchElementException as e:
        print(f'Não Carregou a página: {e}')

carregamento()

reports = "logistic-operator/reports"
page_reports = base_url + reports
driver.get(page_reports)

def download_relatorio():
    try:
        # Abrir opções de relatório
        driver.find_element(By.XPATH, '//*[@id="downshift-2-toggle-button"]').click()

        time.sleep(1)

        # Escolher Opção Perfomance
        report_perfomance = driver.find_element(By.XPATH, '//*[@id="downshift-2-item-0"]')
        report_perfomance.click()

        # Escolher Data inicio relatório
        driver.find_element(By.XPATH, '//*[@id="initialDate"]').click()
        date_start = driver.find_element(By.XPATH, f"//div[@class='DayPicker-inner-day' and text()={last_day}]")
        date_start.click()

        time.sleep(1)

        # Escolher data final do relatório
        driver.find_element(By.XPATH, '//*[@id="finalDate"]').click()
        date_final = driver.find_element(By.XPATH, f"//div[@class='DayPicker-inner-day' and text()={last_day}]")
        date_final.click()

        time.sleep(1)

        # Gerar relatório
        driver.find_element(By.XPATH, '//*[@id="root"]/div[2]/main/div/div/button/div').click()

    except NoSuchElementException as e:
        print(f"Erro: {e}")

time.sleep(1)

download_relatorio()

time.sleep(5)

# Descompactar ZIP com relatório
def unzip(zip_path, extract_to):
    try:
        if not os.path.exists(zip_path):
            raise FileNotFoundError(f"Arquivo Zip não encontrado: {zip_path}")
        
        os.makedirs(extract_to, exist_ok=True)


        with zp.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
            print(f"Arquivos extraídos para: {extract_to}")
    except PermissionError as e:
        print(f"Erro ao descompactar: {e}")

zip_path = r'C:\\Users\\Anderson Luiz\\Downloads\\bundle.zip'
extract_to = r'./backup_relatorio'
unzip(zip_path, extract_to)

time.sleep(2)

# Concatenar relatórios e salvar em um arquivo xls


print()
input("Verifique se o Download foi executado...")

driver.quit()



