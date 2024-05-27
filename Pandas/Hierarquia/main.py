from globo_automacoes.decoradores import main_, retry
from globo_automacoes.logger import logger_setup
from libs.config import Config
from assets.sharepoint_assets import sharepoint_functions_ctx as spf
from src.Valida16caracteres import Hierarquia16Caracteres
from src.Valida13caracteres import Hierarquia13Caracteres

config = Config()
 
log = logger_setup(config.tag_processo, config.metricas.id_execucao)

@main_(config, log)
def main() -> None:
    print('funcionou')
    sharepoint_obj = spf.SharePointFunctions_ctx(site_name="InterfacesRCA")

    sharepoint_obj.download_all_files_from_folder(f"Documentos%20Partilhados/FJI/FinancasCorporativas/HierarquizacaoDeProjetos/HierarquiaDeProjetos", "files")
    sharepoint_obj.download_all_files_from_folder(f"Documentos%20Partilhados/FJI/FinancasCorporativas/HierarquizacaoDeProjetos/PlanilhaDeCarga", "files")
    
    hierarquia16_obj=Hierarquia16Caracteres('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/PROJETO CARGA.xlsm', 'GL_SEGMENT_VALUES_INTERFACE', 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE')
    hierarquia13_obj=Hierarquia13Caracteres('C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/PROJETO CARGA.xlsm', 'GL_SEGMENT_VALUES_INTERFACE', 'C:/Users/madis/Documents/DocPython/Pandas/Hierarquia/files/PP.OO.9.100 - Hierarquia de Projetos.xlsm', 'GL_SEGMENT_HIER_INTERFACE')
    try:
        
        hierarquia16_obj.print_data_frame()
        hierarquia13_obj.print_data_frame()
    except Exception:
        print('falhou')



if __name__ == "__main__":
 
    main()