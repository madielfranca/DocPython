from globo_automacoes.decoradores import retry
from src.LogStatus import LogStatus
from openpyxl import load_workbook
from datetime import datetime, timedelta
import pandas as pd
import logging
import os

log = logging.getLogger(__name__)

class Remove_arquivo_log_dia_anterior:
    def __init__(self) -> None:
        self.valor_carga = self


    @retry(3, Exception)
    def remove_arquivo_log_dia_anterior(self):
        valor_carga = self.valor_carga
        file_name_carga = self.valor_carga
        current_date = datetime.now().strftime('%Y-%m-%d')
        print(valor_carga)
        print(file_name_carga)
        log = logging.getLogger(__name__)

        data_atual = datetime.now()

        # Calcular o dia anterior
        previous_day = data_atual - timedelta(days=1)
        previous_day.strftime("%Y-%m-%d")

        file_path = f'HierarquiaStatus_{previous_day.strftime("%Y-%m-%d")}.xlsx'

        if os.path.exists(file_path):

            log.info("Removendo Arquivo log dia anterior %s", file_path)
            os.remove(file_path)