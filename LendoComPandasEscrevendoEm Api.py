import sys
import os
import glob
import pandas as pd
import json
import time
from openpyxl import load_workbook
from datetime import datetime
import pyperclip
from io import StringIO
str_io_obj = StringIO()
import requests
import re
import warnings
import shutil

sys.path.append('..\\..\\..\\Comum\\Scripts\\oracle_cloud_assets')
from get_azure_api_token import GetAzureApiToken

# DESATIVANDO WARNINGS
warnings.filterwarnings('ignore')

def teste_api_158(token):
	url = "https://"

	# CAPTURA DE USR DA MAQUINA
	profile = os.environ["USERPROFILE"]
	pathSharepoint = '{usr}\\{sharepoint}'.format(usr=profile,
												  sharepoint='General\\900Sustentacao\\60Robos\\P4355_Criação_e_Carga_de_Requisicoes_Internas_no_ERP')

	path = glob.glob(pathSharepoint + '\\Input\\*.xlsm')

	for file in path:

		# DEFININDO DATAFRAME
		dic_dtype = {'ID': str, 'Agrupador de RI': str, 'Data necessidade': str, }
		df_Requisicoes = pd.read_excel(file, sheet_name='Solicitação', skiprows=8, dtype=dic_dtype)

		# LIMPANDO VALORES NAN
		df_Requisicoes = df_Requisicoes.dropna(subset=['ID'])

		df_concatenados = df_Requisicoes[['Agrupador de RI', 'Descrição da RI', 'E-mail solicitante', 'E-mail CC']]

		df_concatenados['Status'] = ''
		df_concatenados['RI'] = ''
		df_concatenados['Id'] = ''

		df_Requisicoes['Centro de Custo'] = df_Requisicoes['Centro de Custo'].astype(int)
		df_Requisicoes['Centro de Resultado'] = df_Requisicoes['Centro de Resultado'].astype(int)
		df_Requisicoes['Projeto'] = df_Requisicoes['Projeto'].astype(int)
		df_Requisicoes['Finalidade'] = df_Requisicoes['Finalidade'].astype(int)

		df_Requisicoes.rename(columns={0: 'number',
									   'Item*': 'itemNumber',
									   'ID': 'id',
									   'Organização de inventário': 'destinationOrganizationCode',
									   'Subinventário*': 'warehouseId',
									   'Data necessidade': 'requestDeliveryDate',
									   # 'Quantidade*':'quantity',
									   'UM': 'primaryUnitOfMeasureCode',
									   'E-mail solicitante': 'requesterEmail',
									   'Retirada no balcão': 'pickupAtLocation',
									   'Fornecedores CNPJ': 'processingSupplierCode',

									   }, inplace=True)

		df_Requisicoes['Centro de Custo'][df_Requisicoes['Projeto ou Chave'] != 'Chave'] = ''
		df_Requisicoes['Centro de Resultado'][df_Requisicoes['Projeto ou Chave'] != 'Chave'] = ''
		df_Requisicoes['Projeto'][df_Requisicoes['Projeto ou Chave'] != 'Chave'] = ''
		df_Requisicoes['Finalidade'][df_Requisicoes['Projeto ou Chave'] != 'Chave'] = ''

		# SEPARANDO CODIGO DO ITEM
		# itemNumber = df_Requisicoes['itemNumber'].str.split(' ',expand=True)
		# df_Requisicoes = df_Requisicoes.assign(itemNumber=itemNumber[0].str.strip())

		taskDescription = df_Requisicoes['Tarefa'].str.split(' ', expand=True)
		df_Requisicoes = df_Requisicoes.assign(taskDescription=taskDescription[1].str.strip())

		taskNumber = df_Requisicoes['Tarefa'].str.split(' ', expand=True)
		df_Requisicoes = df_Requisicoes.assign(taskNumber=taskNumber[0].str.strip())

		projectDecription = df_Requisicoes['Projeto PPM'].str.split('|', expand=True)
		df_Requisicoes = df_Requisicoes.assign(projectDecription=projectDecription[1].str.strip())

		projectNumber = df_Requisicoes['Projeto PPM'].str.split('|', expand=True)
		df_Requisicoes = df_Requisicoes.assign(projectNumber=projectNumber[0].str.strip())

		# SEPARANDO RESTANTE DAS INFORMAÇÕES
		df_RI = df_Requisicoes[['id', 'Agrupador de RI']]
		df_RI.rename(columns={'id': 'LinhaExcel'}, inplace=True)
		df_RI.drop_duplicates(subset=['Agrupador de RI'], inplace=True)
		df_RI.reset_index(drop=True, inplace=True)

		# AGRUPANDO CAMPO "LIST" DA API EM UM CONJUNTO

		wb = load_workbook(filename=file)
		sheet_ranges = wb["Solicitação"]
		prepareEmail = sheet_ranges['B3'].value

		df_Requisicoes.insert(1, 'sourceOrderNumber', 'REQLIN00001006')
		# df_Requisicoes.insert(1, 'sourceOrderNumber', df_Requisicoes['id'])
		df_Requisicoes.insert(5, 'sourceOrganizationCode', df_Requisicoes['destinationOrganizationCode'])
		df_Requisicoes['id'] = ''
		df_Requisicoes.insert(9, 'quantity', 1)
		df_Requisicoes.insert(10, 'prepareEmail', prepareEmail)
		df_Requisicoes.insert(12, 'destinationTypeCode', 'EXPENSE')
		df_Requisicoes.insert(13, 'TranferOrderLineDestination', '300000061722416')

		df_Requisicoes['requestDeliveryDate'] = df_Requisicoes['requestDeliveryDate'].str.replace('00:00:00', '')
		df_Requisicoes['requestDeliveryDate'] = df_Requisicoes['requestDeliveryDate'].str.strip()

		df_Processado = (df_Requisicoes.groupby(['Agrupador de RI'], as_index=False)
						 .apply(lambda x: x[['sourceOrderNumber',
											 'itemNumber',
											 'destinationOrganizationCode',
											 'sourceOrganizationCode',
											 'warehouseId',
											 'requestDeliveryDate',
											 'quantity',
											 'primaryUnitOfMeasureCode',
											 'prepareEmail',
											 'requesterEmail',
											 'destinationTypeCode',
											 'TranferOrderLineDestination',
											 'Tarefa',
											 'taskDescription',
											 'projectDecription',
											 'taskNumber',
											 'projectNumber',
											 'Tipo de Dispêndio',
											 'Projeto ou Chave',
											 'Centro de Custo',
											 'Centro de Resultado',
											 'Projeto PPM',
											 'Projeto',
											 'Finalidade',
											 ]]
								.to_dict('r')))

		df_Processado = df_Processado.to_frame()
		df_Processado.rename(columns={0: 'lines'}, inplace=True)
		df_Processado.reset_index(drop=True, inplace=True)

		# INSERINDO NOVOS CAMPOS DO CABEÇALHO
		df_Processado.insert(0, 'status', 'NEW')
		data_atual = str(datetime.today())
		data_atual = data_atual[:-16]
		sourceOrder = re.sub('[\-\:\" "]', '', data_atual)
		df_Processado.insert(1, 'date', data_atual)
		# df_Processado.insert(2, 'sourceOrderNumber', str(sourceOrder) + '' + df_RI['Agrupador de RI'])
		df_Processado.insert(2, 'sourceOrderNumber', 'REQHDR00000006')
		df_Processado.insert(3, 'pickupAtLocation', 'S')
		# df_Processado.insert(3, 'pickupAtLocation',df_Requisicoes['pickupAtLocation'])
		df_Processado.insert(4, 'addressComplement', 'RI sem Projeto')

		# MONTANDO O CAMPO "LINES"
		for i in df_Processado.index:
			for j in range(len(df_Processado['lines'][i])):

				'''SE O CAMPO "PROJETO OU CHAVE" TIVER O VALOR "CHAVE", PREENCHER O CAMPO DE ACCOUNTSEGMENTS.
                CASO CONTRATIO, PREENCHER O CAMPO DISTIBUTIONS'''
				if df_Processado['lines'][i][j]['Projeto ou Chave'] != 'Chave':
					df_Processado['lines'][i][j]['distribution'] = (
						[
							{
								'number': 1,
								'quantity': 1,
								# 'quantity':df_Processado['lines'][i][j]['quantity'],
								"accountingSegments": [
									{
										"type": "finalidade/motivo",
										"value": df_Processado['lines'][i][j]['Finalidade'],
									},

								]

							}
						]
					)
				else:
					df_Processado['lines'][i][j]['distribution'] = (
						[
							{
								'number': 1010,
								'quantity': 1,
								"accountingSegments": [
									{
										"type": "centroDeResultado",
										# "value": df_Processado['lines'][i][j]['Centro de Resultado'],
										"value": '0000000',
									},
									{
										"type": "projeto",
										"value": '0000000000000000',
										# "value": df_Processado['lines'][i][j]['Projeto'],
									},
									{
										"type": "finalidade",
										"value": 'FG000006',
										# "value": df_Processado['lines'][i][j]['Finalidade'],
									},
									{
										"type": "centroDeCusto",
										"value": 'GL200408002',
										# "value": df_Processado['lines'][i][j]['Centro de Custo'],
									},
								]
							}
						]
					)
				# REMOVENDO CHAVES DESNECESSÁRIAS
				del df_Processado['lines'][i][j]['Centro de Custo']
				del df_Processado['lines'][i][j]['Centro de Resultado']
				del df_Processado['lines'][i][j]['Projeto PPM']
				del df_Processado['lines'][i][j]['Projeto']
				del df_Processado['lines'][i][j]['Finalidade']
				del df_Processado['lines'][i][j]['taskNumber']
				del df_Processado['lines'][i][j]['projectDecription']
				del df_Processado['lines'][i][j]['projectNumber']
				del df_Processado['lines'][i][j]['Tarefa']
				del df_Processado['lines'][i][j]['taskDescription']
				del df_Processado['lines'][i][j]['Tipo de Dispêndio']
				del df_Processado['lines'][i][j]['Projeto ou Chave']

		# FAZENDO O POST POR REQUISIÇÃO

		for ind in df_Processado.index:
			data = df_Processado.iloc[ind:ind + 1].to_json(orient='records')

			data = json.loads(data)
			print(json.dumps(data[0], indent=2, ensure_ascii=False))
			data = json.dumps(data[0], ensure_ascii=False)
	headers = {
		'Content-Type': 'application/json',
		'x-api-key': '',
		'Authorization': 'Bearer {token}'.format(token=token)
	}
	response = requests.request("POST", url, headers=headers, data=data)

	print(data)
	print(response)


def test_api_159():
	api_url = "https://"
	grant_type = "client_credentials"
	client_id = ""
	client_secret = ""
	resource = ""

	# CAPTURANDO O TOKEN
	get_token_obj = GetAzureApiToken(api_url, grant_type, client_id, client_secret, resource)
	token = get_token_obj.get_api_token()

	url = ""

	payload={}
	headers = {
	  'x-api-key': '',
	  'Authorization': 'Bearer {token}'.format(token=token[1])
	}

	response = requests.request("GET", url, headers=headers, data=payload)

	print(response.text)

def main():
	try:

		api_url = "https://login.microsoftonline.com//oauth2/token"
		grant_type = "client_credentials"
		client_id = ""
		client_secret = ""
		resource = ""


		# CAPTURANDO O TOKEN
		get_token_obj = GetAzureApiToken(api_url, grant_type, client_id, client_secret, resource)
		token = get_token_obj.get_api_token()

		teste_api_158(token[1])
		test_api_159()






	except Exception as e:
		print(str(e))
		time.sleep(20)
		raise e

if __name__ == "__main__":
	main()
