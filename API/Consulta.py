import requests
import json
import time
import pandas as pd
import pyperclip

print('Envio')

input_path = r'logProcessados\logProcessadosJprs.xlsx'
file_path = r"../../../Comum/Scripts/batch_python_executor/ParametroPython.txt"

def P5225_techcosts_log_erro():
  #LENDO PARAMETROS DO AUTOMATION
  with open(file_path, "r") as file:
      content = file.read()
      values = content.split(',')
      numDocumento = values[0]
      cnpjFornecedor = values[1]
    
  #LENDO ARQUIVO DE LOG    
  data_frame = pd.read_excel(input_path)

  try:
    for indice, linha in data_frame.iterrows():
        if str(linha['numDocumento']) == str(numDocumento):
          values = linha['horaProcesso'].split('T')
          LogId = linha['LogId']
          Protocolo = linha['Protocolo']
          Descricao = linha['Descrição']
          data = values[0]
          hora = values[1]
          horaLimpa = hora.split('.')
          horaProcesso = data + " " + horaLimpa[0]
          if str(LogId) == 'nan':
            linha['LogId'] = 'null'
          if str(Protocolo) == 'nan':
            linha['Protocolo'] = 'null'
          if str(Descricao) == 'nan':
            linha['Descrição'] = 'null'


          #ENVIANDO DADOS DE LOG
          url = "https://api.4mapit.com.br/clientes/globo/jprs/pagamentos"

          payload = json.dumps({
            "dados": [
              {
                "horaProcesso": horaProcesso,
                "id": linha['Id'],
                "numDocumento": linha['numDocumento'],
                "cnpjFornecedor": linha['cnpjFornecedor'],
                "descricao": linha['Descrição'],
                "status": linha['status'],
                "logId": linha['LogId'],
                "protocolo": linha['Protocolo']
              },
            ]
          })
          headers = {
            'authClientKey': 'rgaUjGaL5iAwpNSw1EafHlNE1Gp1f24M4kbEPff3eOs=',
            'authClientSecret': 'o+3gQFAbA0gCbCY8FNss6BFIXCuhdUN+SIJNHiP0zP4=',
            'Content-Type': 'application/json',
            'Cookie': 'connect.sid=s%3AyfmsgDvscx_tJi9gULna1O3RFW6U_Xxy.KNt%2Fj2naBZPSz4MKFmZCaoTgnLcwUbPsdujWEr2RZ5I'
          }

          response = requests.request("POST", url, headers=headers, data=payload)

          print(response.text)
        
          print('______________')

          time.sleep(3)

          #OBTENDO RESPOSTA DOS DADOS CADASTRADOS NA API
          print('Resposta')

          url = f"https://api.4mapit.com.br/clientes/globo/jprs/pagamentos?numDocumento={numDocumento}"

          payload = ""
          headers = {
            'authClientKey': 'rgaUjGaL5iAwpNSw1EafHlNE1Gp1f24M4kbEPff3eOs=',
            'authClientSecret': 'o+3gQFAbA0gCbCY8FNss6BFIXCuhdUN+SIJNHiP0zP4=',
            'Cookie': 'connect.sid=s%3AyfmsgDvscx_tJi9gULna1O3RFW6U_Xxy.KNt%2Fj2naBZPSz4MKFmZCaoTgnLcwUbPsdujWEr2RZ5I'
          }

          response = requests.request("GET", url, headers=headers, data=payload)

          df = pd.DataFrame(response)

          print(response.text)
          print('______________')

          result = 'OK'
          pyperclip.copy(result)

  except:
      result = 'NOK'
      pyperclip.copy(result)
      raise Exception(response.text)

if (__name__ == "__main__"):
     P5225_techcosts_log_erro()

  