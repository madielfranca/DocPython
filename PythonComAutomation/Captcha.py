import csv
import pyperclip
import cv2

# importing os module
import os

# Image path
# image_path = r'D:\gitlab-madielaadm\Automation Anywhere\Bots\Imoveis\TaxaDeIncendio\Docs\Input\Boleto_de_pagamento\temp\Captcha\image.png'

# Image directory
# directory = r'D:\gitlab-madielaadm\Automation Anywhere\Bots\Imoveis\TaxaDeIncendio\Docs\Input\Boleto_de_pagamento\temp\Captcha'

def read_input(image_path, directory):

    try:

        # Using cv2.imread() method
        # to read the image
        img = cv2.imread(image_path)
        gray = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
        thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

        # Remove horizontal
        horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (25,1))
        detected_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
        cnts = cv2.findContours(detected_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        cnts = cnts[0] if len(cnts) == 2 else cnts[1]
        for c in cnts:
            cv2.drawContours(img, [c], -1, (255,255,255), 2)

        # Repair image
        repair_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1,6))
        result = 255 - cv2.morphologyEx(255 - img, cv2.MORPH_CLOSE, repair_kernel, iterations=1)

        # Change the current directory
        # to specified directory
        os.chdir(directory)

        # List files and directories

        print("Before saving image:")
        print(os.listdir(directory))

        # Filename
        filename = 'savedImage.jpg'

        # Using cv2.imwrite() method
        # Saving the image
        cv2.imwrite(filename, thresh)

        # List files and directories

        print("After saving image:")
        print(os.listdir(directory))

        print('Successfully saved')
        pyperclip.copy("OK")
        
    except Exception as ex:
        pyperclip.copy("NOK")
        error = f"Falha ao gerar relatório. Erro: {ex}"
        raise Exception(error)

if (__name__ == "__main__"):

    # Recebe o input o caminho onde o csv do relatório será salvo
    # Esse campo é oriundo da task do AA P4345_Orquestrador
    input_requisicoes = pyperclip.paste()

    with open('countries.csv', 'w', encoding='UTF8') as f:
        writer = csv.writer(f)
        # write the data
        writer.writerow(input_requisicoes)

    # Remove qualquer lixo que seja originado no processo de copiar
    input_requisicoes = input_requisicoes.replace(r'\r\n', '').strip()
    image_path, directory = input_requisicoes.split(";")
    print(input_requisicoes)
read_input(image_path, directory)