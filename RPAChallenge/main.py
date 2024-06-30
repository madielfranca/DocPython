# main.py

from modules.InputForms import Form
import os

# Configura o User da maquina
user_path = os.path.expanduser("~")

# URL do formulário
url_formulario = 'https://rpachallenge.com'

# Caminho para o arquivo Excel
caminho_planilha = f'{user_path}/Documents/DocPython/RPAChallenge/files/challenge.xlsx'

def main():
    
    person = Form(url_formulario, caminho_planilha)
    person.inputForms()


if __name__ == "__main__":
    main()
