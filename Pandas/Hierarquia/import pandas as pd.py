import pandas as pd

df_carga = pd.read_excel('C:/Users/madis/Documents/Trampo-docs/Hierarquia/PROJETO CARGA.xlsm', sheet_name='GL_SEGMENT_VALUES_INTERFACE')
df_carga = df_carga.drop(index=[0, 1, 2])

df_hierarquia = pd.read_excel('C:/Users/madis/Documents/Trampo-docs/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', sheet_name='GL_SEGMENT_HIER_INTERFACE', header=0)
df_hierarquia = df_hierarquia.drop(index=[0, 1])



# df_hierarquia.rename(columns={0: 'A',
# 									   '*Value Set Code': 'Unnamed: 0',
# 									   'ID': 'id',
# 									   'Organização de inventário': 'destinationOrganizationCode',
# 									   'Subinventário*': 'warehouseId',
# 									   'Data necessidade*': 'requestDeliveryDate',
# 									   'Quantidade*':'quantity',
# 									   'UM': 'primaryUnitOfMeasureCode',
# 									   'E-mail solicitante': 'requesterEmail',
# 									   'Retirada no balcão': 'pickupAtLocation',
# 									   'Fornecedores CNPJ': 'processingSupplierCode',
# 									   'Organization ID': 'TranferOrderLineDestination',
# 									   'Endereço complementar': 'EnderecoComplementar'
# 									   }, inplace=True)
print(df_hierarquia)


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
            quantidade_caracteres = len(valores_projeto_carga)
            print(quantidade_caracteres)
            print(valores_projeto_carga)
            print(valores_hierarquia)
            print('--------------------------------')

nome_arquivo = 'HierarquiaDeProjetos.xlsx'
df_hierarquia.to_excel(nome_arquivo, index=False, header=None )


            

       
