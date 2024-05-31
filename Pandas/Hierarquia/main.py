from assets.sharepoint_assets import sharepoint_functions_ctx as spf
from src.Valida10caracteres import Hierarquia10Caracteres
from src.Valida12caracteres import Hierarquia12Caracteres
from src.Valida13caracteres import Hierarquia13Caracteres
from src.Valida16caracteres import Hierarquia16Caracteres
from globo_automacoes.decoradores import main_, retry
from globo_automacoes.logger import logger_setup
from libs.config import Config
from datetime import datetime
import shutil
import os

config = Config()
 
log = logger_setup(config.tag_processo, config.metricas.id_execucao)

# Get the current date
current_date = datetime.now().strftime('%Y-%m-%d')

@main_(config, log)
def main() -> None:
    sharepoint_obj = spf.SharePointFunctions_ctx(site_name="InterfacesRCA")

    sharepoint_obj.download_all_files_from_folder(f"Documentos%20Partilhados/FJI/FinancasCorporativas/HierarquizacaoDeProjetos/HierarquiaDeProjetos", "files")
    sharepoint_obj.download_all_files_from_folder(f"Documentos%20Partilhados/FJI/FinancasCorporativas/HierarquizacaoDeProjetos/PlanilhaDeCarga", "files")
 
    # Step 1: Specify the source file and destination directory

    source_file = 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/PP.OO.9.100 - Hierarquia de Projetos.xlsm'
    destination_directory = 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia'

    # Ensure the destination directory exists
    os.makedirs(destination_directory, exist_ok=True)

    # Step 2: Construct the full destination path
    destination_path = os.path.join(destination_directory, os.path.basename(source_file))

    # Step 3: Copy the file to the destination directory
    shutil.move(source_file, destination_path)
    print(f'Copied {source_file} to {destination_path}')

    # Step 1: Specify the directory and file pattern
    root_folder  = 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/'
    # file_pattern = '*.xlsm'  # Example pattern to match all .txt files
    # Loop through all files in the root folder and its subfolders
    for folder_path, _, filenames in os.walk(root_folder):
        for filename in filenames:
            # Process the file
            print(os.path.join(folder_path, filename))
            if 'Hierarquia' in filename :
                break
            else:
                hierarquia10_obj=Hierarquia10Caracteres('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/'+filename, 'GL_SEGMENT_VALUES_INTERFACE', 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE', filename)
                hierarquia12_obj=Hierarquia12Caracteres('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/'+filename, 'GL_SEGMENT_VALUES_INTERFACE', 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE', filename)
                hierarquia13_obj=Hierarquia13Caracteres('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/'+filename, 'GL_SEGMENT_VALUES_INTERFACE', 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE', filename)
                hierarquia16_obj=Hierarquia16Caracteres('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/'+filename, 'GL_SEGMENT_VALUES_INTERFACE', 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE', filename)
                try:        
                    hierarquia10_obj.print_data_frame()
                    hierarquia12_obj.print_data_frame()
                    hierarquia13_obj.print_data_frame()
                    hierarquia16_obj.print_data_frame()
                except Exception:
                    print('falhou')
    arquivo_log = f'HierarquiaStatus_{current_date}.xlsx'
    sharepoint_obj.upload_file(f"Documentos%20Partilhados/FJI/FinancasCorporativas/HierarquizacaoDeProjetos/HierarquiaDeProjetos", "C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/Hierarquia", "PP.OO.9.100 - Hierarquia de Projetos.xlsm")
    sharepoint_obj.upload_file(f"Documentos%20Partilhados/FJI/FinancasCorporativas/HierarquizacaoDeProjetos/StatusExecucaoRobo", "C:/Users/madis/Documents/DocPython/Pandas/Hierarquia", arquivo_log)

if __name__ == "__main__":
 
    main()