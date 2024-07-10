from assets.sharepoint_assets import sharepoint_functions_ctx as spf
from src.P5494_ValidarProjetosSemHierarquia import Validar_projetos_sem_hierarquia
from src.P5494_RemoveArquivoLogDiaAnterior import Remove_arquivo_log_dia_anterior
from src.P5494_MoveArquivosSharePoint import Move_arquivos_SharePoint
from src.P5494_Valida10caracteres import Hierarquia10Caracteres
from src.P5494_Valida12caracteres import Hierarquia12Caracteres
from src.P5494_Valida13caracteres import Hierarquia13Caracteres
from src.P5494_Valida16caracteres import Hierarquia16Caracteres
from src.P5494_FormatarFBDI import FormatarFBDI
from globo_automacoes.decoradores import main_, retry
from globo_automacoes.logger import logger_setup
from globo_automacoes.emails import AutomacoesEmail
from datetime import datetime, timedelta
from src.LogStatus import LogStatus
from libs.config import Config
import pandas as pd
import numpy as np
import logging
import shutil
import os

config = Config()
 
current_date = datetime.now().strftime('%Y-%m-%d')
user_path = os.path.expanduser("~")

@main_(config)
def main() -> None:
     
    log = logging.getLogger(__name__)
    obj_email = AutomacoesEmail()

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
                validar_projetos_sem_hierarquia_obj=Validar_projetos_sem_hierarquia()
                remove_arquivo_log_dia_anterior_obj=Remove_arquivo_log_dia_anterior()
                move_arquivos_SharePoint_obj=Move_arquivos_SharePoint()
                FormatarFBDI_obj=FormatarFBDI('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia', 'GL_SEGMENT_HIER_INTERFACE', 'PP.OO.9.100 - Hierarquia de Projetos', '.xlsm')

                try:

                    hierarquia10_obj.validar_projetos_10_caracteres()
                    hierarquia12_obj.validar_projetos_12_caracteres()
                    hierarquia13_obj.validar_projetos_13_caracteres()
                    hierarquia16_obj.validar_projetos_16_caracteres()

                except Exception:
                    #log.info("Hierarquia 10 caracteres executado com sucesso %s", variavel_aqui)
                    obj_email.enviar_email()
                    log.error()
                    print('falhou')

    validar_projetos_sem_hierarquia_obj.validar_projetos_sem_hierarquia()
    remove_arquivo_log_dia_anterior_obj.remove_arquivo_log_dia_anterior()     
    move_arquivos_SharePoint_obj.move_arquivos_SharePoint()
    FormatarFBDI_obj.formatarFBDI()

            
if __name__ == "__main__":
    main()
 