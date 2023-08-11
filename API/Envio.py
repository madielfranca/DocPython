from pathlib import Path
import sys
import requests
import json
import pandas
import os
import pyperclip
import os.path as path
sys.path.append(os.path.join(os.path.dirname(__file__), "../../../Comum/Scripts"))
sys.path.append('..\\..\\..\\Comum\\Scripts\\oracle_cloud_assets')
import log_assets.log_functions as lf
import oracle_cloud_assets.get_azure_api_token as azure_token
from get_azure_api_token import GetAzureApiToken
import cryptography_assets.cryptography_functions as crypto_functions
import setup_assets.setup_functions as sf


def P5225_techcosts():
    #TAG DO SUBPROCESSO A SER UTILIZADO 
    process_tag = "P5225"

    work_dir = path.abspath(path.join(__file__ ,".."))
    dict_setup = sf.load_setup(work_dir + "/" + process_tag + "_config_python.yml")

    work_dir = path.abspath(path.join(__file__ ,".."))
    dict_setup = sf.load_setup(work_dir + "/" + process_tag + "_config_python.yml")


    # SETANDO CREDENCIAIS

    api_url = dict_setup['api_url']
    grant_type = dict_setup['grant_type']
    client_id = dict_setup['client_id']
    client_secret = dict_setup['client_secret']
    resource = dict_setup['resource']

    # CAPTURANDO O TOKEN
    get_token_obj = GetAzureApiToken(api_url, grant_type, client_id, client_secret, resource)
    token = get_token_obj.get_api_token()

    headers = {
        'Content-Type': 'application/json; charset=utf-8',
        'x-api-key': 'n22ddf7zw9xkfbcwnaer9kxk',
        'Authorization': 'Bearer {token}'.format(token=token[1])
    }

    #CONVERTENDO RESULTADO DA REQUISISÃO EM PANDAS
    response = requests.get("""LInk_API""", headers=headers)
    json_data = json.loads(response.text)
    df = pandas.DataFrame(json_data)

    print(df)
    print('Baixando anexos')

    #BAIXANDO ARQUIVOS PDF
    try:
        for row in df.index:
            id = df['id'][row]
            invoiceNumber = df['invoiceNumber'][row]
            supplierCode = df['supplierCode'][row]
            nomeNota = Path('./anexo/notaFiscal.pdf')
            nomeFatura = Path('./anexo/fatura.pdf')

            nota = requests.get(
                f"LInk_API",
                headers=headers)

            fatura = requests.get(
                f"LInk_API",
                headers=headers)

            responseNota = nota
            nomeNota.write_bytes(responseNota.content)

            responseFatura = fatura
            nomeFatura.write_bytes(responseFatura.content)

            os.rename('./anexo/notaFiscal.pdf', f'./anexo/NotaFiscal-{id}-{invoiceNumber}-{supplierCode}.pdf')
            os.rename('./anexo/fatura.pdf', f'./anexo/Fatura-{id}-{invoiceNumber}-{supplierCode}.pdf')

        df.to_csv('./temp/Dados_TechCosts.csv', mode='a', index=False, header=False)

        result = 'OK'
        pyperclip.copy(result)

    except:

        result = 'NOK'
        pyperclip.copy(result)
        raise Exception(json_data)



if (__name__ == "__main__"):
     P5225_techcosts()


