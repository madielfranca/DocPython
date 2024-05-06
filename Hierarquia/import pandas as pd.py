import pandas as pd
import numpy as np

df_carga = pd.read_excel('C:/Users/madis/Documents/Trampo-docs/Hierarquia/PROJETO CARGA.xlsm', sheet_name='GL_SEGMENT_VALUES_INTERFACE')
df_carga = df_carga.drop(index=[0, 1, 2])

df_hierarquia = pd.read_excel('C:/Users/madis/Documents/Trampo-docs/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', sheet_name='GL_SEGMENT_HIER_INTERFACE')
df_hierarquia = df_hierarquia.drop(index=[0, 1, 2])


for row in df_carga.index:
    # # Verifica se o nome da coluna é 'Unnamed: 0'
    if 'Unnamed: 0' in df_carga.columns:
        valores_projeto_carga = df_carga['Unnamed: 1'][row]
        print(valores_projeto_carga)


    else:
        valores_projeto_carga = df_carga['B'][row]
        # print(valores_projeto_carga)


    for row in df_hierarquia.index:
        # # Verifica se o nome da coluna é 'Unnamed: 0'
        if 'Unnamed: 0' in df_hierarquia.columns:
            valores_hierarquia = df_hierarquia['Unnamed: 33'][row]
            # print(valores_hierarquia)
            # print('2')

        else:
            valores_hierarquia = df_hierarquia['AJ'][row]
            # print(valores_hierarquia)


        valores_hierarquia = str(valores_hierarquia)
        valores_projeto_carga = str(valores_projeto_carga)
        nova_string = valores_projeto_carga[:-3]

       
        if valores_hierarquia == nova_string:
            print(valores_projeto_carga)
            print(valores_hierarquia)
            print('--------------------------------')


            

       
