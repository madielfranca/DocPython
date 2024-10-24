import os
import sys
import shutil
import traceback
import io
sys.path.append(os.path.join(os.path.dirname(__file__), '..', ''))
import cryptography_assets.cryptography_functions  as cf
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
class SharePointFunctions_ctx(object):
    """		
	Classe para interface com o sharepoint. Utiliza como base via composição a biblioteca Office365-REST-Python-Client.
	A URL: "https://tvglobocorp.sharepoint.com/sites/ " junto com o site_name será a URL que será utilizada para o acesso. 
    Exemplo de site_name:
        site_name: rpaglobo = https://tvglobocorp.sharepoint.com/sites/rpaglobo
        site_name: interfacesrca = https://tvglobocorp.sharepoint.com/sites/interfacesrca    
        site_name: afiliadas = https://tvglobocorp.sharepoint.com/sites/afiliadas.contratos
    
    Caso o site_name nao seja informado, sera utilizado o site default "RPAGlobo".
    Caso necessario, poderá também informar o usuario e a senha.    
	:ivar username: Usuário de e-mail utilizado para login no Sharepoint. Se nao informado, usara o RCA_USROFFICE@G.GLOBO.  Default: None. 
	:type username: string
	:ivar password: Senha para login no Sharepoint. Se nao informado, usara a senha do usuario RCA_USROFFICE@G.GLOBO Default: None
	:type password: string
	:ivar site_name: Nome do site Sharepoint.
	:type site_name: string
	"""
    def __init__(self, username = None, password = None, site_name='RPAGlobo') -> None:
        if (username == None or password == None):            
            username = "RCA_USROFFICE@G.GLOBO"
            if os.environ.get('RCA_USROFFICE'):
                password = os.environ.get('RCA_USROFFICE')
            else:
                credential_path = os.environ['USERPROFILE'] + "/Globo Comunicação e Participações sa/RPA Globo - Documentos/General/900Sustentacao/80Credenciais"
                key_file = credential_path + "/Keys/SECRET_RCA_USROFFICE.key"
                encrypted_message_file = credential_path + "/Credentials/CRD_RCA_USROFFICE.txt"
                
                password= cf.decrypt_file(encrypted_message_file,key_file)         
        
        self.site_url = f"https://tvglobocorp.sharepoint.com/sites/{site_name}/"
        self._authenticate(username, password)
        
    def _authenticate(self, username, password):
        """
        Autentica o usuário para acessar o contexto do cliente.
        Args:
        - username (str): O nome de usuário para autenticação.
        - password (str): A senha correspondente ao nome de usuário.
        Returns:
        - None
        Esta função cria um contexto de cliente para acesso a url do sharepoint, utilizando as credenciais fornecidas.
        """
        self.ctx: any = ClientContext(self.site_url).with_credentials(UserCredential(username, password))
    def create_folder(self, sharepoint_folder, new_folder_name):
        """
        Cria uma nova pasta no SharePoint.
        Args:
        - sharepoint_folder (str): O caminho da pasta no SharePoint onde a nova pasta será criada.
        Exemplo: "Documentos%20Compartilhados/General/900Sustentacao/60Robos/P4310_PCLDMovimentacaoContabil/Docs/Procs/Output" 
        - new_folder_name (str): O nome da nova pasta a ser criada.
        Returns: 
        - None
        Este método cria uma nova pasta no SharePoint. Verifica se a pasta já existe no caminho fornecido.
        Se não existir, cria a nova pasta dentro do caminho especificado.
        """   
        folder = (
            self.ctx.web.get_folder_by_server_relative_url(f"{sharepoint_folder}{new_folder_name}")
            .select(["Exists"])
            .get()
            .execute_query()
        )
        if folder.exists:
            pass        
        else:
            list_obj = self.ctx.web.get_folder_by_server_relative_url(sharepoint_folder)
            list_obj.folders.add(new_folder_name).execute_query()
   
    def delete_folder(self, sharepoint_folder):
        """
        Exclui uma pasta do SharePoint.
        Args:
        - sharepoint_folder (str): O caminho da pasta no SharePoint que será excluída.
        Exemplo: "Documentos%20Compartilhados/General/900Sustentacao/60Robos/P4310_PCLDMovimentacaoContabil/Docs/Procs/Output" 
        Returns:
        - None
        Este método exclui a pasta especificada no caminho do SharePoint fornecido.
        """
        
        file = self.ctx.web.get_folder_by_server_relative_url(sharepoint_folder)
        file.delete_object().execute_query()
 
    def download_file(self, sharepoint_folder, destination_path, file_name):
        """
        Baixa um arquivo do SharePoint para um diretório local.
        Args:
        - sharepoint_folder (str): O caminho da pasta no SharePoint onde o arquivo está localizado.
        Exemplo: "Documentos%20Compartilhados/General/900Sustentacao/60Robos/P4310_PCLDMovimentacaoContabil/Docs/Procs/Output" 
        - destination_path (str): O caminho do diretório local onde o arquivo será baixado.
        - file_name (str): O nome do arquivo que será baixado.
        Returns:
        - None
        Este método baixa o arquivo especificado do SharePoint para um diretório local. 
        Verifica se o diretório de destino existe e o cria se não existir.
        """       
        # Create local destination_path if it doesn't exist
        if not os.path.exists(destination_path):
            os.makedirs(destination_path)
        download_file_path = f"{destination_path}/{file_name}"
        sharepoint_file_path = f"{sharepoint_folder}/{file_name}"
        with open(download_file_path, "wb") as local_file:
            file = (
                self.ctx.web.get_file_by_server_relative_url(sharepoint_file_path)
                .download(local_file)
                .execute_query()
            ) 
    def upload_file(self, sharepoint_folder, local_path, file_name):
        """
        Faz o upload de um arquivo local para uma pasta no SharePoint.
        Args:
        - sharepoint_folder (str): O caminho da pasta no SharePoint onde o arquivo será enviado.
        Exemplo: "Documentos%20Compartilhados/General/900Sustentacao/60Robos/P4310_PCLDMovimentacaoContabil/Docs/Procs/Output" 
        - local_path (str): O caminho local onde o arquivo está localizado.
        - file_name (str): O nome do arquivo que será enviado.
        Returns:
        - str: Retorna "OK" se o arquivo for enviado com sucesso.
        Este método realiza o upload de um arquivo local para uma pasta específica no SharePoint.
        """
                   
        upload_file_path = f"{local_path}/{file_name}"
        folder = self.ctx.web.get_folder_by_server_relative_url(sharepoint_folder)
        with open(upload_file_path, mode='rb') as file:
            file = folder.files.upload(file).execute_query()
    def upload_all_files_from_folder(self, sharepoint_folder, local_path):
        """
        Faz o upload de todos os arquivos de uma pasta local para uma pasta no SharePoint.
        Args:
        - sharepoint_folder (str): O caminho da pasta no SharePoint onde o arquivo será enviado.
        Exemplo: "Documentos%20Compartilhados/General/900Sustentacao/60Robos/P4310_PCLDMovimentacaoContabil/Docs/Procs/Output" 
        - local_path (str): O caminho local dos arquivos que serão feitos os uploads.        
        Returns:
        - str: Retorna "OK" se o arquivo for enviado com sucesso.
        Este método realiza o upload de todos os arquivos de uma pasta local para uma pasta específica no SharePoint.
        """
                
        folder = self.ctx.web.get_folder_by_server_relative_url(sharepoint_folder)
        # Iterate through all files in the local folder
        for filename in os.listdir(local_path):
            local_file_path = os.path.join(local_path, filename)
            if os.path.isfile(local_file_path):  # Ensure it's a file (not a directory)
                with open(local_file_path, mode='rb') as file:
                    file = folder.files.upload(file).execute_query()
    def delete_file(self, sharepoint_folder, file_name_w_ext):
        """
        Exclui um arquivo específico de uma pasta no SharePoint.
        Args:
        - sharepoint_folder (str): O caminho da pasta no SharePoint onde o arquivo está localizado.
        Exemplo: "Documentos%20Compartilhados/General/900Sustentacao/60Robos/P4310_PCLDMovimentacaoContabil/Docs/Procs/Output" 
        - file_name_w_ext (str): O nome do arquivo com extensão que será excluído.
        Returns:
        - None
        Este método exclui um arquivo específico de uma pasta no SharePoint, com base no nome do arquivo e no caminho da pasta.
        """
            
        sharepoint_file_relative_path = f"{sharepoint_folder}/{file_name_w_ext}"
        file = self.ctx.web.get_file_by_server_relative_url(sharepoint_file_relative_path)
        file.delete_object().execute_query()
    def delete_all_files_from_folder(self, sharepoint_folder):
        """
        Exclui todos os arquivos de uma pasta no SharePoint.
        Args:
        - sharepoint_folder (str): O caminho da pasta no SharePoint onde o arquivo está localizado.
        Exemplo: "Documentos%20Compartilhados/General/900Sustentacao/60Robos/P4310_PCLDMovimentacaoContabil/Docs/Procs/Output" 
        Returns:
        - None
        Este método exclui todos os arquivos de uma pasta no SharePoint.
        """
        files = self.ctx.web.get_folder_by_server_relative_url(sharepoint_folder).files
        self.ctx.load(files).execute_query()
        
        for file in files:
            file.delete_object().execute_query()
 
    def list_folder_files(self, sharepoint_folder):
        """
        Lista os arquivos em uma pasta específica do SharePoint.
        Args:
        - sharepoint_folder (str): O caminho da pasta no SharePoint a ser listada.
        Exemplo: "Documentos%20Compartilhados/General/900Sustentacao/60Robos/P4310_PCLDMovimentacaoContabil/Docs/Procs/Output" 
        Returns:
        - list: Retorna uma lista contendo os caminhos dos arquivos na pasta especificada.
        Este método retorna uma lista dos caminhos dos arquivos presentes na pasta especificada do SharePoint.
        """
        files = self.ctx.web.get_folder_by_server_relative_url(sharepoint_folder).files
        self.ctx.load(files).execute_query()
        
        sharepoint_files_list = [f.serverRelativeUrl for f in files]
        return sharepoint_files_list
   
    def download_all_files_from_folder (self, sharepoint_folder, destination_path ):
        """
        Faz o download de todos os arquivos de uma pasta do SharePoint para um diretório local.
        Args:
        - sharepoint_folder (str): O caminho da pasta no SharePoint que contém os arquivos a serem baixados.
        Exemplo: "Documentos%20Compartilhados/General/900Sustentacao/60Robos/P4310_PCLDMovimentacaoContabil/Docs/Procs/Output" 
        - destination_path (str): O caminho do diretório local onde os arquivos serão baixados.
        Returns:
        - None
        Este método faz o download de todos os arquivos presentes na pasta especificada do SharePoint para um diretório local.
        """
        files = self.ctx.web.get_folder_by_server_relative_url(sharepoint_folder).files
        self.ctx.load(files).execute_query()
        # Create local subfolder if it doesn't exist
        if not os.path.exists(destination_path):
            os.makedirs(destination_path, exist_ok=True)
        for file in files:
            file_name = file.properties['Name']
            download_file_path = f"{destination_path}/{file_name}"
            with open(download_file_path, "wb") as local_file:
                file.download(local_file).execute_query()
  
    def download_all_files_from_folder_and_subfolders(self, sharepoint_folder, destination_path):
        """
        Baixa todos os arquivos de uma pasta e subpastas do SharePoint para um diretório local.
        Args:
        - sharepoint_folder (str): O caminho da pasta no SharePoint que contém os arquivos a serem baixados.
        Exemplo: "Documentos%20Compartilhados/General/900Sustentacao/60Robos/P4310_PCLDMovimentacaoContabil/Docs/Procs/Output" 
        - destination_path (str): O caminho do diretório local onde os arquivos serão baixados.
        Returns:
        - None
        Este método inicia o processo de download de todos os arquivos presentes na pasta especificada e suas subpastas no SharePoint,
        salvando-os em um diretório local específico.
        """
       
        folder = self.ctx.web.get_folder_by_server_relative_url(sharepoint_folder)
        self.ctx.load(folder).execute_query()
        self._download_files_from_folder_recursively(folder, destination_path)
        
    def _download_files_from_folder_recursively(self, folder, destination_path):
        """
        Método auxiliar para baixar recursivamente arquivos de uma pasta e subpastas do SharePoint.
        Args:
        - folder: O objeto da pasta do SharePoint.
        - destination_path (str): O caminho do diretório local onde os arquivos serão baixados.
        Este método é utilizado internamente para baixar recursivamente todos os arquivos presentes na pasta especificada e suas subpastas no SharePoint,
        salvando-os em um diretório local específico.
        """
        # Get all files in the current folder
        files = folder.files
        self.ctx.load(files)
        self.ctx.execute_query()
        # Create local subfolder if it doesn't exist
        if not os.path.exists(destination_path):
            os.makedirs(destination_path, exist_ok=True)
        # Download files in the current folder
        for file in files:
            file_name = file.properties['Name']
            download_file_path = f"{destination_path}/{file_name}"
            with open(download_file_path, "wb") as local_file:
                file.download(local_file).execute_query()
        # Get all subfolders in the current folder
        subfolders = folder.folders
        self.ctx.load(subfolders)
        self.ctx.execute_query()
        # Recursively download files from subfolders
        for subfolder in subfolders:
            subfolder_destination_path = os.path.join(destination_path, subfolder.properties['Name'])
            
            # Create local subfolder if it doesn't exist
            if not os.path.exists(subfolder_destination_path):
                os.makedirs(subfolder_destination_path, exist_ok=True)
            # Recursively download files from subfolder
            self._download_files_from_folder_recursively(subfolder, subfolder_destination_path)
                
    def get_data_buffer(self, sharepoint_folder, file_name):
        """
        Le o arquivo e o converte em um buffer de dados
        Args:
        - sharepoint_folder (str): O caminho da pasta no SharePoint onde o arquivo está localizado.
        Exemplo: "Documentos%20Compartilhados/General/900Sustentacao/60Robos/P4310_PCLDMovimentacaoContabil/Docs/Procs/Output" 
        - file_name (str): O nome do arquivo que será convertido.
        Returns:
        - buffer de dados: Retorna um buffer com base no conteúdo do arquivo convertido.
        Este método lê um arquivo localizado em uma pasta do SharePoint e o converte em um buffer de dados.
        """
                   
        sharepoint_file_path = f"{sharepoint_folder}/{file_name}"
        
        file = self.ctx.web.get_file_by_server_relative_url(sharepoint_file_path)            
        self.ctx.load(file)
        self.ctx.execute_query()
        file_content  = file.read()
        buffer = io.BytesIO(file_content)
        return buffer
 