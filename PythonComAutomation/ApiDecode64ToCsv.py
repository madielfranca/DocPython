import sys
import os

sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "../../../Comum/Scripts"))
from oracle_cloud_assets.oracle_cloud_reports_rest_api import ReportsRestAPI, ReportHelpfulMethods

reports_params = "[]"

reports_path = "/Custom/Integracoes/Reports/RP_CANCEL_NFSE.xdo"

# Cria o objeto a partir do contruttor da classe
obj_request = ReportsRestAPI(reports_path, reports_params, 'DEV10')


# Executa o método que utiliza a api rest

api_output = obj_request.get_api_return()

if api_output[0] == 'OK':

    # Cria um objeto a partir do contrutor da classe métodos úteis

    report_methods = ReportHelpfulMethods()

    # Decodifica o texto em base64 retornado pela API

    decoded_text = report_methods.decode_base64_reports(api_output[1])

    data = decoded_text.strip()

    if data.count('\n') > 1:

        report_methods.save_report_in_csv('temp.csv', decoded_text)

        result = 'OK'

    else:

        result = 'NOK'

else:

    raise Exception(api_output[1])


