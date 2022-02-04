import os
import glob
import pandas as pd

input_path = "C:\Users\Madiel\Documents"
output_path = "C:\Users\Madiel\Documents\saida"

def read_input(input_path, output_path):
    """
    Função para ler o arquivo de entrada e retornar um dataframe pandas
    """

    # DEFININDO DATAFRAME

    df_Requisicoes = pd.read_excel(input_path, sheet_name='PGTOS  GLOBO RJ', skiprows=5)
    df_concatenados = df_Requisicoes[['CONTA', 'PROJETO', 'FINALIDADE', 'PARCELAS', 'INSCRIÇÃO ', 'SITE']]
    df_concatenados.to_excel(output_path + '\IPTU CONTROLE DE PGTOS.xlsx', index=False)

    return df_concatenados

read_input(input_path, output_path)