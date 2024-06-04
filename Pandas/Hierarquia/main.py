from assets.sharepoint_assets import sharepoint_functions_ctx as spf
from src.Valida10caracteres import Hierarquia10Caracteres
from src.Valida12caracteres import Hierarquia12Caracteres
from src.Valida13caracteres import Hierarquia13Caracteres
from src.Valida16caracteres import Hierarquia16Caracteres
from globo_automacoes.decoradores import main_, retry
from globo_automacoes.logger import logger_setup
from datetime import datetime, timedelta
from src.LogStatus import LogStatus
from libs.config import Config
import pandas as pd
import os
import shutil
import math

config = Config()
 
log = logger_setup(config.tag_processo, config.metricas.id_execucao)

# Get the current date
current_date = datetime.now().strftime('%Y-%m-%d')

@main_(config, log)
def main() -> None:
    sharepoint_obj = spf.SharePointFunctions_ctx(site_name="InterfacesRCA")

    sharepoint_obj.download_all_files_from_folder(f"Documentos%20Partilhados/FJI/FinancasCorporativas/HierarquizacaoDeProjetos/HierarquiaDeProjetos", "files")
    sharepoint_obj.download_all_files_from_folder(f"Documentos%20Partilhados/FJI/FinancasCorporativas/HierarquizacaoDeProjetos/PlanilhaDeCarga", "files")
 
    # Step 1: Specify the source file and destination directory

    source_file = 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/PP.OO.9.100 - Hierarquia de Projetos.xlsm'
    destination_directory = 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia'

    # Ensure the destination directory exists
    os.makedirs(destination_directory, exist_ok=True)

    # Step 2: Construct the full destination path
    destination_path = os.path.join(destination_directory, os.path.basename(source_file))

    # Step 3: Copy the file to the destination directory
    shutil.move(source_file, destination_path)
    print(f'Copied {source_file} to {destination_path}')

    # Step 1: Specify the directory and file pattern
    root_folder  = 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/'
    # file_pattern = '*.xlsm'  # Example pattern to match all .txt files
    # Loop through all files in the root folder and its subfolders
    for folder_path, _, filenames in os.walk(root_folder):
        for filename in filenames:
            # Process the file
            print(os.path.join(folder_path, filename))
            if 'Hierarquia' in filename :
                break
            else:
                hierarquia10_obj=Hierarquia10Caracteres('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/'+filename, 'GL_SEGMENT_VALUES_INTERFACE', 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE', filename)
                hierarquia12_obj=Hierarquia12Caracteres('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/'+filename, 'GL_SEGMENT_VALUES_INTERFACE', 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE', filename)
                hierarquia13_obj=Hierarquia13Caracteres('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/'+filename, 'GL_SEGMENT_VALUES_INTERFACE', 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE', filename)
                hierarquia16_obj=Hierarquia16Caracteres('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/'+filename, 'GL_SEGMENT_VALUES_INTERFACE', 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE', filename)
                try:        
                    hierarquia10_obj.validar_projetos_10_caracteres()
                    hierarquia12_obj.validar_projetos_12_caracteres()
                    hierarquia13_obj.validar_projetos_13_caracteres()
                    hierarquia16_obj.validar_projetos_16_caracteres()
                except Exception:
                    print('falhou')
    arquivo_log = f'HierarquiaStatus_{current_date}.xlsx'
    # sharepoint_obj.upload_file(f"Documentos%20Partilhados/FJI/FinancasCorporativas/HierarquizacaoDeProjetos/HierarquiaDeProjetos", "C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia", "PP.OO.9.100 - Hierarquia de Projetos.xlsm")
    sharepoint_obj.upload_file(f"Documentos%20Partilhados/FJI/FinancasCorporativas/HierarquizacaoDeProjetos/StatusExecucaoRobo", "C:/Users/madis/Documents/DocPython/Pandas/Hierarquia", arquivo_log)


def limpa_arquivo_log_dia_anterior():

    # Get the current date
    data_atual = datetime.now()
    print(data_atual)
    # Calculate the previous day
    previous_day = data_atual - timedelta(days=1)
    previous_day.strftime("%Y-%m-%d")

    file_path = f'HierarquiaStatus_{previous_day.strftime("%Y-%m-%d")}.xlsx'
    print(file_path)
    if os.path.exists(file_path):
        os.remove(file_path)

def find_value_in_files():
    values_to_add  = []
    df_status = pd.read_excel('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/HierarquiaStatus_2024-06-05.xlsx', engine='openpyxl')

    caminho_pasta   = 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/'

    # Lista para armazenar os DataFrames
    lista_dfs = []

    # Loop para ler cada arquivo Excel na pasta
    for arquivo in os.listdir(caminho_pasta):
        if arquivo.endswith('.xlsm'):  # Certifique-se de que é um arquivo Excel
            caminho_arquivo = os.path.join(caminho_pasta, arquivo)
            df = pd.read_excel(caminho_arquivo, 'GL_SEGMENT_VALUES_INTERFACE')
            lista_dfs.append(df)

    # Combina todos os DataFrames em um único DataFrame
    df_combinado = pd.concat(lista_dfs, ignore_index=True)


    for row in df_combinado.index:
        valor_carga = df_combinado['Unnamed: 1'][row]
        if not (valor_carga is None or (isinstance(valor_carga, float) and math.isnan(valor_carga))):
            if valor_carga != '*Value':
                encontrado = False
                for row in df_status.index:
                    valor_status = df_status['Status'][row]
                    nome_status = df_status['Arquivo'][row]

                    if valor_carga in valor_status:
                        print(f"Value found in file: {valor_carga}")
                        encontrado = True
                   
                if encontrado == False:
                    print(f"Value not found in file: {valor_carga}")
                    values_to_add.append("Erro na hierarquia."+ valor_carga)
                 
                    break   
    LogStatus_obj=LogStatus(values_to_add, nome_status)
    LogStatus_obj.logar()
            
if __name__ == "__main__":
 
    main()
    limpa_arquivo_log_dia_anterior()
    find_value_in_files()
