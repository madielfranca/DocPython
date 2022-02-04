import os
import glob
import pandas as pd

# profile = os.environ["USERPROFILE"]
# pathSharepoint = '{usr}\\{sharepoint}'.format(usr=profile,
#                                               sharepoint='Globo Comunicação e Participações sa\\RPA Globo - Documentos\\General\\900Sustentacao\\60Robos\\P4345_ImoveisGeraçãodeGuiasIPTU')
#
# # input_path = glob.glob(pathSharepoint + '\\Input\\iptu.xlsx')

input_path = "D:\gitlab-madielaadm\Automation Anywhere\Bots\Imoveis\P4375_GeraçãoDeGuiasIPTU\Docs\Input\controle_de_pagamento\modelo\IPTU CONTROLE DE PGTOS 2022 (RJ).xlsm"
output_path = "D:\gitlab-madielaadm\Automation Anywhere\Bots\Imoveis\P4375_GeraçãoDeGuiasIPTU\Docs\Input\controle_de_pagamento"

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