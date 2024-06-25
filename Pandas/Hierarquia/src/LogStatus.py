from globo_automacoes.decoradores import retry
from datetime import datetime
import pandas as pd
import logging
import os


class LogStatus:
    def __init__(self, values_to_add, file_name_caraga) -> None:
        self.valor_carga = values_to_add
        self.file_name_carga = file_name_caraga

    @retry(3, Exception)
    def logar(self):
        valor_carga = self.valor_carga
        current_date = datetime.now().strftime('%Y-%m-%d')
        log_date = datetime.now()

        # Configura o log
        logging.basicConfig(filename=f'app_{current_date}.log', filemode='a', format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)

        try:
            file_path = f'HierarquiaStatus_{current_date}.xlsx'

            if os.path.exists(file_path):
                existing_df = pd.read_excel(file_path)
            else:
                # Se o arquivo não existir, cria um DataFrame vazio com nomes de colunas especificados
                existing_df = pd.DataFrame(columns=['Status'])

            values_to_add_list = valor_carga
            # Cria um DataFrame com os novos valores

            new_data = pd.DataFrame({'Status': values_to_add_list, 'Arquivo': self.file_name_carga, 'Data Execução': log_date})

            # Anexa os novos dados ao DataFrame existente
            existing_df = pd.concat([existing_df, new_data], ignore_index=True)

            # Grava o DataFrame atualizado de volta no arquivo Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                existing_df.to_excel(writer, index=False, sheet_name='Sheet1')

        except Exception as e:
            logging.error(f'Error occurred: {e}')

        logging.info('Script finished.')