from globo_automacoes.sharepoint import SharepointContext, SharepointFile
from sistema.task import Sistema, SistemaLoginInput
# from shared.webdriver import obter_driver_local
from globo_automacoes.base_config import Config
from globo_automacoes.decoradores import main_
from arquivei.db_arquivei import SQLServerDb
from dotenv import load_dotenv, find_dotenv
from oracle.api_oracle import ApiOracle
from config import Config, obter_config
import xml.etree.ElementTree as ET
import pandas as pd
import requests
import base64
import json
import os
 

load_dotenv(find_dotenv())
CONFIG: Config = obter_config() # type: ignore
# CONFIG: Config = Config()

# Leitura do arquivo CSV
files = 'files/output/RPT46 - Relatorio documentos fiscais V3_RPT46 - Relatorio documentos fiscais V3.csv'

df = pd.read_csv(files)

# Filtrar valores que não são "CAMPO RESERVADO PARA AUDITOR"
chaves = df[df['DANFE_NF'] != "CAMPO RESERVADO PARA AUDITOR"]['DANFE_NF']

# Exibir apenas os valores filtrados da coluna 'danfe_nf'
print(chaves)

# Acessar as variáveis de ambiente
SQLServerDb_obj=SQLServerDb()
SQLServerDb_obj.conectar_database()
lote_tamanho = 100
resultados = []

# Dividir as chaves em lotes de 100
for i in range(0, len(chaves), lote_tamanho):
    lote = chaves[i:i + lote_tamanho]
    chaves_formatadas = ", ".join([f"'{chave}'" for chave in lote])

    # Query para cada lote de chaves
    query = f"SELECT * FROM TB_RPA_NFE_ARQUIVEI WHERE CHAVE_ACCESSO IN ({chaves_formatadas})"
    # query = f"SELECT XML_ANX FROM TB_RPA_NFE_ARQUIVEI WHERE CHAVE_ACCESSO IN ({chaves_formatadas})"
    resultado_lote = SQLServerDb_obj.select_chamado(query)
    for result in resultado_lote:
        xml_str = result[0]
        # input(xml_str)

        # Agregar os resultados
        resultados.extend(resultado_lote)



SQLServerDb_obj.create_xml(resultados)
print('SQLServerDb_obj')


# Caminho do diretório onde está o arquivo XML

caminho_arquivo = 'notas_fiscais.xml'
print(caminho_arquivo)



import xml.etree.ElementTree as ET
import pandas as pd

# Função auxiliar para extrair o texto de um elemento XML
def extract_text(element):
    return element.text if element is not None else None

# Função para processar o arquivo XML e extrair os dados
def process_nfe_file(caminho_arquivo, output_excel):
    try:
        tree = ET.parse(caminho_arquivo)
        root = tree.getroot()

        notas_fiscais = []

        # Percorre as notas fiscais
        for nota_fiscal in root.findall('NotaFiscal'):
            xml_data = nota_fiscal.find('XML').text # type: ignore

            # Analisa o XML da nota fiscal
            nfe_tree = ET.ElementTree(ET.fromstring(xml_data)) # type: ignore
            nfe_root = nfe_tree.getroot()
            inf_nfe = nfe_root.find('.//infNFe')
            if inf_nfe is not None:
                # Extração dos dados da nota fiscal
                nfe_data = {
                    'Número NF-e': extract_text(nfe_root.find('.//nNF')),
                    'Série': extract_text(nfe_root.find('.//serie')),
                    'CNPJ Emitente': extract_text(nfe_root.find('.//emit//CNPJ')),
                    'CNPJ Destinatário': extract_text(nfe_root.find('.//dest//CNPJ')),
                    'CPF Destinatário': extract_text(nfe_root.find('.//dest//CPF')),
                    'Data de Emissão': extract_text(nfe_root.find('.//dhEmi')),
                    'Valor Frete': extract_text(nfe_root.find('.//vFrete')),
                    'Valor Seguro': extract_text(nfe_root.find('.//vSeg')),
                    'Valor Desconto': extract_text(nfe_root.find('.//vDesc')),
                    'Valor Outros': extract_text(nfe_root.find('.//vOutro')),
                    'Valor Total': extract_text(nfe_root.find('.//vNF')),
                    'Valor Produtos': extract_text(nfe_root.find('.//vProd')),
                    'Código Produto': extract_text(nfe_root.find('.//cProd')),
                    'Descrição Produto': extract_text(nfe_root.find('.//xProd')),
                    'Unidade Comercial': extract_text(nfe_root.find('.//uCom')),
                    'NCM': extract_text(nfe_root.find('.//NCM')),
                    'Quantidade': extract_text(nfe_root.find('.//qCom')),
                    'Valor Unitário': extract_text(nfe_root.find('.//vUnCom')),
                    'CFOP': extract_text(nfe_root.find('.//CFOP')),
                    'Natureza Operação': extract_text(nfe_root.find('.//natOp')),
                    'Referência NF-e': extract_text(nfe_root.find('.//refNFe')),
                    'Informações Fisco': extract_text(nfe_root.find('.//infAdFisco')),
                    'Informações Complementares': extract_text(nfe_root.find('.//infCpl')),
                    'CEST': extract_text(nfe_root.find('.//CEST')),
                    'ICMS': extract_text(nfe_root.find('.//vICMS')),
                    'Base Cálculo ST': extract_text(nfe_root.find('.//vBCST')),
                    'ICMS ST': extract_text(nfe_root.find('.//pICMSST')),
                    'CST': extract_text(nfe_root.find('.//CST')),
                    'Nome Emitente': extract_text(nfe_root.find('.//emit//xNome')),
                    'Logradouro': extract_text(nfe_root.find('.//enderEmit//xLgr')),
                    'Número': extract_text(nfe_root.find('.//enderEmit//nro')),
                    'Complemento': extract_text(nfe_root.find('.//enderEmit//xCpl')),
                    'Bairro': extract_text(nfe_root.find('.//enderEmit//xBairro')),
                    'Município': extract_text(nfe_root.find('.//enderEmit//xMun')),
                    'UF': extract_text(nfe_root.find('.//enderEmit//UF')),
                    'CEP': extract_text(nfe_root.find('.//enderEmit//CEP')),
                    'País': extract_text(nfe_root.find('.//enderEmit//xPais')),
                    'Inscrição Estadual': extract_text(nfe_root.find('.//IE')),
                    'Consumidor Final': extract_text(nfe_root.find('.//indFinal')),
                    'Origem': extract_text(nfe_root.find('.//orig')),
                    'Valor IPI': extract_text(nfe_root.find('.//vIPI')),
                    'Valor COFINS': extract_text(nfe_root.find('.//vCOFINS')),
                    'Valor PIS': extract_text(nfe_root.find('.//vPIS')),
                    'Valor II': extract_text(nfe_root.find('.//vII')),
                    'Valor FCP': extract_text(nfe_root.find('.//vFCP')),
                    'Valor FCP ST': extract_text(nfe_root.find('.//vFCPST')),
                    'Valor FCP ST Retido': extract_text(nfe_root.find('.//vFCPSTRet'))
                }

                # Adiciona os dados da nota fiscal à lista
                notas_fiscais.append(nfe_data)

        # Criar um DataFrame a partir da lista de notas fiscais
        df = pd.DataFrame(notas_fiscais)
        print(df)
        # Salvar em um arquivo Excel
        df.to_excel(output_excel, index=False)

        print(f"Planilha '{output_excel}' criada com sucesso!")

    except FileNotFoundError:
        print(f"Arquivo {caminho_arquivo} não encontrado no diretório.")
    except ET.ParseError:
        print("Erro ao analisar o arquivo XML.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

# Exemplo de uso
caminho_arquivo = 'notas_fiscais.xml'
output_excel = 'saida_nfe.xlsx'
process_nfe_file(caminho_arquivo, output_excel)


# breakpoint()

# @main_(config=CONFIG)
# def main() -> None:
#     """Executor da Automação"""
#     print(CONFIG.area)
#     breakpoint()
    

#     # URL da API
#     url = "https://oic-prod-gruszg9zlxj2-gr.integration.ocp.oraclecloud.com/ic/api/integration/v1/flows/rest/ERP_REPORTDOWNLOAD_UPSERT/1.0/report"
#     # url = "https://oic-uat-erp-gruszg9zlxj2-gr.integration.ocp.oraclecloud.com/ic/api/integration/v1/flows/rest/ERP_REPORTDOWNLOAD_UPSERT/1.0/report"

#     # Payload que você está enviando
#     payload = json.dumps({
#         "filePath": "Custom/RELATORIOS AJUSTADOS/RELATORIOS/RPT46 - Relatorio documentos fiscais V3.xdo",
#         "filters": [
#             {
#                 "value": "2024-09-15T21:00:00.000-03:00",
#                 "key": "DATA_INICIO"
#             },
#                 {
#                 "value": "DANFE",
#                 "key": "TIPO_DOC"
#             },
#             {
#                 "value": "2024-09-24T21:00:00.000-03:00",
#                 "key": "DATA_FIM"
#             }       
            
#         ]
#     })
#     print(type(payload))
#     print(payload)
#     # Cabeçalhos com autenticação
#     headers = {
#         'Content-Type': 'application/json',
#         'Authorization': 'Basic T0lDX1BSRF9FUlBfUlBBX1NFUlZJQ0VfVVNFUl9CQVNJQ0FVVEg6NjJjODQ0ZTItNmY4MC00MTBlLWE0ODQtZmFmM2ZhOTU2ZTll'
#     }

#     # Fazendo a requisição
#     response = requests.post(url, headers=headers, data=payload)
#     response.raise_for_status()
#     # Verifica o status da resposta

#     print("Requisição bem-sucedida!")

#     # Decodificando o conteúdo Base64 do campo 'report'
#     dados_report = response.json()
#     print(dados_report)

#     file_base64 = dados_report['report']

#     # Tentar salvar o arquivo como csv, já que o conteúdo parece ser um PDF
#     with open("relatorio_documentos_fiscais.csv", "wb") as f:
#         f.write(base64.b64decode(file_base64))

#     print("Arquivo csv salvo com sucesso como 'relatorio_documentos_fiscais.csv'.") 

    
 
   
  
  
    

# if __name__ == '__main__':
#     main()
