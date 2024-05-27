from globo_automacoes.decoradores import retry
from datetime import datetime
import pandas as pd
import logging
import os


class LogStatus:
    def __init__(self, values_to_add) -> None:
        self.valor_carga = values_to_add

    @retry(3, Exception)
    def logar(self):
        valor_carga = self.valor_carga

        # Get the current date
        current_date = datetime.now().strftime('%Y-%m-%d')

        # Configure logging
        logging.basicConfig(filename=f'app_{current_date}.log', filemode='a', format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)

        # Log the start of the script
        logging.info('Script started.')

        try:
            # Path to the existing Excel file
            file_path = f'HierarquiaStatus_{current_date}.xlsx'

            file_name = os.path.basename(file_path)
            
            # Split the file name by underscore and extract the last part
            date_part = file_name.split("_")[-1]

            dia_atual = datetime.now().day
            dia_anterior = date_part.split("-")[-1].split(".")[0]
            dia_anterior = int(dia_anterior)

            if dia_atual > dia_anterior:
                os.remove(file_path)
            
            # Check if the file exists
            if os.path.exists(file_path):
                # Read the existing Excel file
                existing_df = pd.read_excel(file_path)
                logging.info(f'File {file_path} found and loaded.')
            else:
                # If the file does not exist, create an empty DataFrame with specified column names
                existing_df = pd.DataFrame(columns=['Status'])
                logging.info(f'File {file_path} not found. Created a new DataFrame.')

            # Example list of values to add
            values_to_add_list = valor_carga
            # Create a DataFrame with the new values
            new_data = pd.DataFrame({'Status': values_to_add_list})

            # Append the new data to the existing DataFrame
            existing_df = pd.concat([existing_df, new_data], ignore_index=True)
            logging.info(f'Added new values to the DataFrame.')

            # Write the updated DataFrame back to the Excel file
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                existing_df.to_excel(writer, index=False, sheet_name='Sheet1')
                logging.info(f'Data written to {file_path}.')

        except Exception as e:
            logging.error(f'Error occurred: {e}')

        logging.info('Script finished.')
