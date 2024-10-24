import csv
import pyperclip
import pandas as pd


def read_input(input_path, output_path):
    """
    Função para ler o arquivo de entrada e retornar um dataframe pandas
    """

    try:
        # DEFININDO DATAFRAME

        df_Requisicoes = pd.read_excel(input_path, sheet_name='PGTOS  GLOBO', skiprows=8)
        df_concatenados = df_Requisicoes[['PROJETO', 'FINALIDADE', 'IMÓVEL', 'CBMERJ', 'SITE']]
        df_concatenados.to_excel(output_path + '\IPTU CONTROLE DE PGTOS.xlsx', index=False)
        pyperclip.copy("OK")

    except Exception as ex:
        pyperclip.copy("NOK")
        error = f"Falha ao gerar relatório. Erro: {ex}"
        raise Exception(error)

if (__name__ == "__main__"):

    # Recebe o input o caminho onde o csv do relatório será salvo
    # Esse campo é oriundo da task do AA P4345_Orquestrador
    input_requisicoes = pyperclip.paste()

    # Remove qualquer lixo que seja originado no processo de copiar
    input_requisicoes = input_requisicoes.replace(r'\r\n', '').strip()
    input_path, output_path = input_requisicoes.split(";")






read_input(input_path, output_path)
