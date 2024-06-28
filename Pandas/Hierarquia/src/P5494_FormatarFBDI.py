from assets.sharepoint_assets import sharepoint_functions_ctx as spf
from globo_automacoes.decoradores import retry
from src.LogStatus import LogStatus
from openpyxl import load_workbook
from datetime import datetime, timedelta
from tqdm import tqdm  # For displaying progress bar
from libs.config import Config
import pandas as pd
import logging
import shutil
import os

log = logging.getLogger(__name__)
current_date = datetime.now().strftime('%Y-%m-%d')
user_path = os.path.expanduser("~")

config = Config()

class FormatarFBDI:
    def __init__(self, pathFBDI, SheetName, WorkbookName, ext) -> None:
        self.pathFBDI = pathFBDI
        self.SheetName = SheetName
        self.WorkbookName = WorkbookName
        self.ext = ext

    @retry(3, Exception)
    def formatarFBDI(self):
        pathFBDI = self.pathFBDI
        SheetName = self.SheetName
        WorkbookName = self.WorkbookName
        ext = self.ext
        log = logging.getLogger(__name__)
        log.info("Formatando arquivo FBDI")
        read_file = pd.read_excel(pathFBDI + "\\" + WorkbookName + ext, sheet_name=SheetName)
        df = pd.DataFrame(read_file)
        numRows = len(read_file)
        Parent = [""] * 36
        maxParent = 99
        lastParentCol = 36

        for i in tqdm(range(3, numRows)):

            if (df.iloc[i - 1, 2] != df.iloc[i, 2]) or (df.iloc[i - 1, 1] != df.iloc[i, 1]) or (df.iloc[i - 1, 0] != df.iloc[i, 0]):
                Parent = [""] * 36
                maxParent = 99
                lastParent = ""
            
            for j in range(35, 5, -1):
                if pd.notna(df.iloc[i, j]) and df.iloc[i, j] != "":
                    df.at[i, 36] = df.iloc[i, j]

                    if j == 35 and lastParent:
                        df.at[i, 37] = lastParent
                    elif Parent[j - 1] and lastParentCol >= j - 1:
                        df.at[i, 37] = Parent[j - 1]
                    elif maxParent == j or maxParent == 99:
                        df.at[i, 37] = "None"
                    else:
                        print(f"Wrong Hierarchy for value at row {i + 1}. Please validate hierarchy and correct the errors.")
                        df = df.drop(columns=df.columns[36:39])
                        return -1
                    
                    Parent[j] = df.iloc[i, j]
                    if j < maxParent:
                        maxParent = j
                    if j != 35:
                        lastParent = df.iloc[i, j]
                        lastParentCol = j
                    break

        numRows = len(df)

        # Percorrer as linhas de 3 a numRows
        for i in range(3, numRows):
            flag = 0
            # Percorra as colunas de 35 a 5 (etapa -1)
            for j in range(35, 5, -1):
                # Verifique se o valor da célula não está vazio
                if pd.notna(df.iloc[i, j]):
                    # Atribuir valor à célula no índice (i, 39)
                    df.at[i, 39] = j - 4
                    float_number = df.at[i, 39]
                    string_number = str(float_number)
                    # Remove '.0' se presente
                    if string_number.endswith('.0'):
                        string_number = string_number[:-2]

                    df.at[i, 39] = string_number
                    flag = 1
                    break
            # Se o sinalizador ainda for 0, saia do loop
            if flag == 0:
                break

        df.drop(df.index[[0, 1, 2]], axis=0, inplace=True)

        # Definir o intervalo de colunas que deseja excluir
        start_column = 5  # Índice da primeira coluna a ser excluída
        end_column = 35   # Índice da última coluna a ser excluída (inclusive)
        log.info("Zipando arquivo FBDI") 
        # Excluir as colunas dentro do intervalo especificado
        df = df.drop(df.columns[start_column:end_column+1], axis=1)

        zip_name = 'GlSegmentHierInterface'


        df.to_csv(pathFBDI + "\\" + zip_name + '.csv', index=False, header=None)
        print(pathFBDI+"\\"+zip_name, 'zip', pathFBDI, zip_name+'.csv')
        shutil.make_archive(pathFBDI+"\\"+zip_name, 'zip', pathFBDI, zip_name+'.csv')