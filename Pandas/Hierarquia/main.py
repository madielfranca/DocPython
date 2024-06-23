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
from tqdm import tqdm  # For displaying progress bar
import pandas as pd
import numpy as np
import shutil
import math
import os

config = Config()
 
log = logger_setup(config.tag_processo, config.metricas.id_execucao)
current_date = datetime.now().strftime('%Y-%m-%d')
user_path = os.path.expanduser("~")

@main_(config, log)
def main() -> None:
    sharepoint_obj = spf.SharePointFunctions_ctx(site_name="InterfacesRCA")
    sharepoint = config.sharepoint_config_file_path
    source_file = config.local_config_source_file_path
    destination_directory = config.local_config_destination_file_path
    destination_directory_status_file = config.local_config_destination_status_file_path

    source_file = f'{user_path}/{source_file}'
    destination_directory = f'{user_path}/{destination_directory}'
    destination_directory_status_file = f'{user_path}/{destination_directory_status_file}'

    arquivos_carga = os.listdir(source_file)

    # Itera sobre cada item na pasta
    for item in arquivos_carga:
        caminho_completo = os.path.join(source_file, item)
        # Verifica se é um arquivo e remove
        if os.path.isfile(caminho_completo):
            os.remove(caminho_completo)
        # Verifica se é um diretório e remove
        elif os.path.isdir(caminho_completo):
            shutil.rmtree(caminho_completo)
    #Baixa os arquivos necessarios do Share Point
    sharepoint_obj.download_all_files_from_folder(f"{sharepoint}/HierarquiaDeProjetos", "files")
    sharepoint_obj.download_all_files_from_folder(f"{sharepoint}/PlanilhaDeCarga", "files")

    os.makedirs(destination_directory, exist_ok=True)

    destination_path = os.path.join(destination_directory, os.path.basename(f'{source_file}/PP.OO.9.100 - Hierarquia de Projetos.xlsm'))
    
    shutil.move(f'{source_file}/PP.OO.9.100 - Hierarquia de Projetos.xlsm', destination_path)
    print(f'Copied {f'{source_file}/PP.OO.9.100 - Hierarquia de Projetos.xlsm'} to {destination_path}')

    # Percorra todos os arquivos na pasta raiz e suas subpastas
    for folder_path, _, filenames in os.walk(source_file):
        for filename in filenames:
            # Process the file
            if 'Hierarquia' in filename :
                break
            else:
                hierarquia10_obj=Hierarquia10Caracteres(folder_path+filename, 'GL_SEGMENT_VALUES_INTERFACE', f'{folder_path}/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE', filename)
                hierarquia12_obj=Hierarquia12Caracteres(folder_path+filename, 'GL_SEGMENT_VALUES_INTERFACE', f'{folder_path}/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE', filename)
                hierarquia13_obj=Hierarquia13Caracteres(folder_path+filename, 'GL_SEGMENT_VALUES_INTERFACE', f'{folder_path}/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE', filename)
                hierarquia16_obj=Hierarquia16Caracteres(folder_path+filename, 'GL_SEGMENT_VALUES_INTERFACE', f'{folder_path}/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE', filename)
                try:        
                    hierarquia10_obj.validar_projetos_10_caracteres()
                    hierarquia12_obj.validar_projetos_12_caracteres()
                    hierarquia13_obj.validar_projetos_13_caracteres()
                    hierarquia16_obj.validar_projetos_16_caracteres()
                except Exception:
                    print('falhou')
    arquivo_log = f'HierarquiaStatus_{current_date}.xlsx'

    sharepoint_obj.upload_file(f"{sharepoint}/StatusExecucaoRobo", destination_directory_status_file, arquivo_log)

def limpa_arquivo_log_dia_anterior():

    data_atual = datetime.now()

    # Calcular o dia anterior
    previous_day = data_atual - timedelta(days=1)
    previous_day.strftime("%Y-%m-%d")

    file_path = f'HierarquiaStatus_{previous_day.strftime("%Y-%m-%d")}.xlsx'

    if os.path.exists(file_path):
        os.remove(file_path)

def validar_projetos_sem_hierarquia():
    values_to_add  = []
    arquivo_log = f'HierarquiaStatus_{current_date}.xlsx'
    df_status = pd.read_excel(f'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/{arquivo_log}', engine='openpyxl')
    caminho_pasta   = 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/'

    # Lista para armazenar os DataFrames
    lista_dfs = []

    # Loop para ler cada arquivo Excel na pasta
    for arquivo in os.listdir(caminho_pasta):
        if arquivo.endswith('.xlsm'):  # Certifique-se de que é um arquivo Excel
            caminho_arquivo = os.path.join(caminho_pasta, arquivo)
            df = pd.read_excel(caminho_arquivo, 'GL_SEGMENT_VALUES_INTERFACE')
            lista_dfs.append(df)

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
                        encontrado = True
                   
                if encontrado == False:
                    values_to_add.append("Erro na hierarquia."+ valor_carga)                
                    break   

    LogStatus_obj=LogStatus(values_to_add, nome_status)
    LogStatus_obj.logar()

def move_arquivos():
    sharepoint_obj = spf.SharePointFunctions_ctx(site_name="InterfacesRCA")
    sharepoint = config.sharepoint_config_file_path
    destination_directory = config.local_config_destination_file_path
    destination_directory = f'{user_path}/{destination_directory}'
    destination_directory_temp = f'{destination_directory}/temp'

    # Especifica o caminho para a pasta
    nome_antigo = 'PP.OO.9.100 - Hierarquia de Projetos.xlsm'
    nome_novo = f'Hierarquia-de-Projetos-{current_date}.xlsm'

    # Caminhos completos dos arquivos
    caminho_antigo = os.path.join(destination_directory, nome_novo)
    caminho_novo = os.path.join(destination_directory_temp, nome_antigo)

    os.makedirs(destination_directory_temp, exist_ok=True)
    shutil.copy(f'{destination_directory}/{nome_antigo}', destination_directory_temp)

    # Verifica se o arquivo existe
    if os.path.isfile(caminho_novo):
        os.rename(caminho_novo, caminho_antigo)

    shutil.move(f'{destination_directory}/{nome_novo}', destination_directory_temp)
 
    # sharepoint_obj.delete_all_files_from_folder(f"{sharepoint}/PlanilhaDeCarga")
    sharepoint_obj.upload_all_files_from_folder(f"{sharepoint}/PlanilhaDeCarga/Processados", "files")
    sharepoint_obj.upload_all_files_from_folder(f"{sharepoint}/HierarquiaDeProjetos/Backup", "files/Hierarquia/temp")

def FormatarFBDI(pathFBDI, SheetName, WorkbookName, ext):
    read_file = pd.read_excel(pathFBDI + "\\" + WorkbookName + ext, sheet_name=SheetName)
    df = pd.DataFrame(read_file)
    numRows = len(read_file)
    Parent = [""] * 36
    maxParent = 99
    lastParentCol = 36

    for i in tqdm(range(3, numRows)):

        if (df.iloc[i - 1, 2] != df.iloc[i, 2]) or (df.iloc[i - 1, 1] != df.iloc[i, 1]) or (df.iloc[i - 1, 0] != df.iloc[i, 0]):
            Parent = [""] * 36
            maxParent = 99
            lastParent = ""
        
        for j in range(35, 5, -1):
            if pd.notna(df.iloc[i, j]) and df.iloc[i, j] != "":
                df.at[i, 36] = df.iloc[i, j]

                if j == 35 and lastParent:
                    df.at[i, 37] = lastParent
                elif Parent[j - 1] and lastParentCol >= j - 1:
                    df.at[i, 37] = Parent[j - 1]
                elif maxParent == j or maxParent == 99:
                    df.at[i, 37] = "None"
                else:
                    print(f"Wrong Hierarchy for value at row {i + 1}. Please validate hierarchy and correct the errors.")
                    df = df.drop(columns=df.columns[36:39])
                    return -1
                
                Parent[j] = df.iloc[i, j]
                if j < maxParent:
                    maxParent = j
                if j != 35:
                    lastParent = df.iloc[i, j]
                    lastParentCol = j
                break

    numRows = len(df)

    # Percorrer as linhas de 3 a numRows
    for i in range(3, numRows):
        flag = 0
        # Percorra as colunas de 35 a 5 (etapa -1)
        for j in range(35, 5, -1):
            # Verifique se o valor da célula não está vazio
            if pd.notna(df.iloc[i, j]):
                # Atribuir valor à célula no índice (i, 39)
                df.at[i, 39] = j - 4
                float_number = df.at[i, 39]
                string_number = str(float_number)
                # Remove '.0' se presente
                if string_number.endswith('.0'):
                    string_number = string_number[:-2]

                df.at[i, 39] = string_number
                flag = 1
                break
        # Se o sinalizador ainda for 0, saia do loop
        if flag == 0:
            break

    df.drop(df.index[[0, 1, 2]], axis=0, inplace=True)

    # Definir o intervalo de colunas que deseja excluir
    start_column = 5  # Índice da primeira coluna a ser excluída
    end_column = 35   # Índice da última coluna a ser excluída (inclusive)

    # Excluir as colunas dentro do intervalo especificado
    df = df.drop(df.columns[start_column:end_column+1], axis=1)

    zip_name = 'GlSegmentHierInterface'

    df.to_csv(pathFBDI + "\\" + zip_name + '.csv', index=False, header=None)
    shutil.make_archive(pathFBDI+"\\"+zip_name, 'zip', pathFBDI, zip_name+'.csv')

            
if __name__ == "__main__":
 
    main()
    limpa_arquivo_log_dia_anterior()
    validar_projetos_sem_hierarquia()
    move_arquivos()
    FormatarFBDI('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia', 'GL_SEGMENT_HIER_INTERFACE', 'PP.OO.9.100 - Hierarquia de Projetos', '.xlsm')