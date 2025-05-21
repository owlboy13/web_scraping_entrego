import time
import os
import zipfile as zp


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