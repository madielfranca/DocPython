# main.py
import subprocess
import pandas as pd

# Executar script1.py
subprocess.run(['python', 'Valida16caracteres.py'])

# Executar script2.py
subprocess.run(['python', 'Valida13caracteres.py'])

# Verificar se os arquivos foram criados
import os

if not os.path.exists('HierarquiaDeProjetos16.xlsx'):
    raise FileNotFoundError('HierarquiaDeProjetos16.xlsx não encontrado.')
if not os.path.exists('HierarquiaDeProjetos13.xlsx'):
    raise FileNotFoundError('HierarquiaDeProjetos13.xlsx não encontrado.')

# Carregar os DataFrames salvos
try:
    df1 = pd.read_excel('HierarquiaDeProjetos16.xlsx')
    df2 = pd.read_excel('HierarquiaDeProjetos13.xlsx')
except Exception as e:
    raise ValueError(f'Erro ao ler os arquivos Excel: {e}')
# Verificar os nomes das colunas
print("Colunas de df1:", df1.columns)
print("Colunas de df2:", df2.columns)

# Selecionar apenas as colunas que precisam ser atualizadas
columns_to_update = ['Parent3', 'Parent2', 'Parent1', 'Value']


df1_selected = df1[columns_to_update]
df2_selected = df2[columns_to_update]

breakpoint()

# Mesclar os DataFrames
merged_df = pd.merge(df1_selected, df2_selected, on='Parent3', suffixes=('_df1', '_df2'))

# Salvar o DataFrame combinado
try:
    merged_df.to_excel('HierarquiaDeProjetos.xlsx', index=False)
except Exception as e:
    raise ValueError(f'Erro ao salvar o arquivo Excel: {e}')

# Visualizar o DataFrame atualizado
print(df1)


