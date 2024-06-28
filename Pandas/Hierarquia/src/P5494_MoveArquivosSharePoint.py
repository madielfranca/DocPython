from assets.sharepoint_assets import sharepoint_functions_ctx as spf
from globo_automacoes.decoradores import retry
from src.LogStatus import LogStatus
from openpyxl import load_workbook
from datetime import datetime, timedelta
from libs.config import Config
import pandas as pd
import logging
import shutil
import os

log = logging.getLogger(__name__)
current_date = datetime.now().strftime('%Y-%m-%d')
user_path = os.path.expanduser("~")

config = Config()

class Move_arquivos_SharePoint:
    def __init__(self) -> None:
        self.valor_carga = self


    @retry(3, Exception)
    def move_arquivos_SharePoint(self):
        log = logging.getLogger(__name__)
        sharepoint_obj = spf.SharePointFunctions_ctx(site_name="InterfacesRCA")
        sharepoint = config.sharepoint_config_file_path
        destination_directory = config.local_config_destination_file_path
        destination_directory = f'{user_path}/{destination_directory}'
        destination_directory_temp = f'{destination_directory}/temp'
        destination_directory_status_file = config.local_config_destination_status_file_path
        destination_directory_status_file = f'{user_path}/{destination_directory_status_file}'
        arquivo_log = f'HierarquiaStatus_{current_date}.xlsx'
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
    
        sharepoint_obj.delete_all_files_from_folder(f"{sharepoint}/PlanilhaDeCarga")
        sharepoint_obj.upload_all_files_from_folder(f"{sharepoint}/PlanilhaDeCarga/Processados", "files")
        sharepoint_obj.upload_all_files_from_folder(f"{sharepoint}/HierarquiaDeProjetos/Backup", "files/Hierarquia/temp")
            
        sharepoint_obj.upload_file(f"{sharepoint}/HierarquiaDeProjetos", "C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia", "PP.OO.9.100 - Hierarquia de Projetos.xlsm")
        sharepoint_obj.upload_file(f"{sharepoint}/StatusExecucaoRobo", destination_directory_status_file, arquivo_log)
        log.info("Movendo arquivos no SharePoint")