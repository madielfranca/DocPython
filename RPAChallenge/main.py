# main.py

from modules.InputForms import Form
import yaml
import os

# Função para carregar o arquivo YAML
def carregar_configuracao(caminho_arquivo):
    with open(caminho_arquivo, 'r') as arquivo:
        configuracao = yaml.safe_load(arquivo)
    return configuracao

# Carregar a configuração do arquivo YAML
configuracao = carregar_configuracao('config_python.yml')

# Acessar caminhos
url_formulario = configuracao['paths']['rpa_challenge_path']
caminho_planilha = configuracao['paths']['input_path']
# log_path = configuracao['paths']['log_path']

# Configura o User da maquina
user_path = os.path.expanduser("~")

def main():  
    obj_inputForms = Form(url_formulario, caminho_planilha)
    obj_inputForms.inputForms()

if __name__ == "__main__":
    main()
