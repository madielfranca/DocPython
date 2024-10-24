from globo_automacoes.decoradores import retry
from src.LogStatus import LogStatus
from openpyxl import load_workbook
from datetime import datetime, timedelta
import pandas as pd
import logging
import math
import os

log = logging.getLogger(__name__)
current_date = datetime.now().strftime('%Y-%m-%d')

class Validar_projetos_sem_hierarquia:
    def __init__(self) -> None:
        self.valor_carga = self

    @retry(3, Exception)
    def validar_projetos_sem_hierarquia(self):
        log = logging.getLogger(__name__)

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
        log.info("Validando projetos sem hierarquia")