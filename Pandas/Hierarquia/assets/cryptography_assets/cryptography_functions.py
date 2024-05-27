from cryptography.fernet import Fernet
import os
import traceback
import os.path as path

def generate_key(secret_key_path = None, service_name = "secret"):
	""" 

	Gera uma chave de segurança e salva no arquivo service_name + ".key" dentro do diretório definido no argumento secret_key_path
	
	:param secret_key_path: Diretório onde o arquivo "secret_key_path + '.key'" será armazenado. Default: SharePoint (RPA Globo - Documentos/General/900Sustentacao/80Credenciais/Keys/)
	:type secret_key_path: string
	:param service_name: Diretório onde o arquivo "secret_key_path + '.key'" será armazenado. Default: "secret"
	:type service_name: string
	
	>>> cryptography_functions.generate_key(secret_key_path = None, service_name = "service-name")

	"""
	if (secret_key_path == None):
		secret_key_path = os.environ['USERPROFILE'] + "\\Globo Comunicação e Participações sa\\RPA Globo - Documentos\\General\\900Sustentacao\\80Credenciais\\Keys\\"
		key_file = "SECRET_" + service_name
	else:
		key_file = "SECRET_" + service_name

	try:
		key = Fernet.generate_key()
		with open(secret_key_path + key_file + ".key", "wb") as key_file:
			key_file.write(key)

	except Exception as e:
		print("Error while generating key...", str(e))
		raise (e)

def load_key(key_file = None, service_name = "secret"):
	"""

	Captura a chave dentro do arquivo service_name + ".key" e retorna seu valor para ser consumido.
	
	:param key_file: caminho para o arquivo ".key" que armazena a chave
	:type key_file: string
	:param service_name: Diretório onde o arquivo "secret_key_path + '.key'" será armazenado. Default: "secret"
	:type service_name: string

	:return: valor da chave de segurança

	>>> cryptography_functions.load_key(key_file = None, service_name = "service-name")

	"""
	if (key_file == None):
		key_file = os.environ['USERPROFILE'] + "\\Globo Comunicação e Participações sa\\RPA Globo - Documentos\\General\\900Sustentacao\\80Credenciais\\Keys\\" + "SECRET_" + service_name + ".key"
	try:
		return open(key_file, "rb").read()

	except Exception as e:
		print("Error while loading key...", str(e))
		raise (e)

def encrypt_message(message, encrypted_message_file = None, key_file = None, service_name = "secret"):
	"""
	
	Criptografa a messagem e grava o resultado no arquivo especificado.
	
	:param encrypted_message_file: Caminho completo e nome do arquivo para armazenar a mensagem criptografada
	:type encrypted_message_file: string
	:param message: Mensagem a ser criptografada
	:type message: string
	:param key_file: Caminho completo e nome do arquivo com a chave de segurança utilizada para criptografar a mensagem
	:type key_file: string

	>>> cryptography_functions.encrypt_message(encrypted_message_file = None, "message", key_file = None, service_name = "service-name")

	"""

	if (encrypted_message_file == None):
		encrypted_message_file = os.environ['USERPROFILE'] + "\\Globo Comunicação e Participações sa\\RPA Globo - Documentos\\General\\900Sustentacao\\80Credenciais\\Credentials\\" + "CRD_" + service_name + ".txt"
	if (key_file == None):
		key_file = os.environ['USERPROFILE'] + "\\Globo Comunicação e Participações sa\\RPA Globo - Documentos\\General\\900Sustentacao\\80Credenciais\\Keys\\" + "SECRET_" + service_name + ".key"
	try:
		key = load_key(key_file)
		encoded_message = message.encode()
		f = Fernet(key)
		encrypted_message = f.encrypt(encoded_message)
		with open(encrypted_message_file, "wb") as encrypted_file:
			encrypted_file.write(encrypted_message)
	
	except Exception as e:
		print("Error while encrypting message...", str(e))
		raise (e)

def decrypt_message(encrypted_message, key_file):
	""" 

	Descriptografa a mensagem.
	
	:param encrypted_message: Mensagem criptografada
	:type encrypted_message: string
	:param key_file: Caminho completo e nome do arquivo com a chave de segurança utilizada para descriptografar a mensagem
	:type key_file: string
	
	:return: Retorna a mensagem descriptografada.

	>>> cryptography_functions.decrypt_message(b'gAAAAABfsspyDs_mXzXkYAkb4YfPgtf0P4-mTMNDiTZKBJwQPjIcNEzmiwK_-K502Y4LFRwMnR3HJ0R13TS0d-9GaTsW0sW9GmG729hueo88QUPaL7WXjo_mYh-HqpqiFLAJv12sWPgtL9puAhyIUd6Z4WQBkWgtIwDUNd-DtEsuqbHKQesqf1SKv2GY3butyXNZu_Dnh-C0', "path to key file")
	
	"""

	try:
		key = load_key(key_file)
		f = Fernet(key)
		decrypted_message = f.decrypt(encrypted_message)

		return decrypted_message.decode()

	except Exception as e:
		print("Error while decrypting message...", str(e))
		raise (e)

def decrypt_file(file_name = None, key_file = None, service_name = "secret"):
	""" 

	Descriptografa a mensagem dentro de um arquivo.
		
	:param file_name: Arquivo com a mensagem criptografada
	:type file_name: string
	:param key_file: Caminho completo e nome do arquivo com a chave de segurança utilizada para descriptografar a mensagem
	:type key_file: string
	
	:return: Retorna a mensagem, contida no arquivo, descriptografada 
	
	"""
	if (file_name == None):
		file_name = os.environ['USERPROFILE'] + "\\Globo Comunicação e Participações sa\\RPA Globo - Documentos\\General\\900Sustentacao\\80Credenciais\\Credentials\\" + "CRD_" + service_name + ".txt"
	if (key_file == None):
		key_file = os.environ['USERPROFILE'] + "\\Globo Comunicação e Participações sa\\RPA Globo - Documentos\\General\\900Sustentacao\\80Credenciais\\Keys\\" + "SECRET_" + service_name + ".key"

	try:
		key = load_key(key_file)
		f = Fernet(key)
		with open(file_name, "rb") as file:
			# read the encrypted data
			encrypted_data = file.read()
		# decrypt data
		#print(encrypted_data)
		decrypted_data = f.decrypt(encrypted_data)
		return str(decrypted_data.decode())

	except Exception as e:
		print("Error while decrypting file...", str(e))
		raise (e)   

if __name__ == "__main__":
	
	scenario = 2
	work_dir = path.abspath(path.join(__file__ ,".."))
	#Cenário 1:
	if(scenario == 1):
		secret_key_path = work_dir + '/tests/'
		print(secret_key_path)
		key_file = secret_key_path + "/SECRET_test_v2.key"
		encrypted_message_file = secret_key_path + "/encrypted_data_test_v2.txt"
		message = "'test message cenário 1'"
		generate_key(secret_key_path, "test_v2")
		encrypt_message(message, encrypted_message_file, key_file)
		print(str(decrypt_file(encrypted_message_file, key_file, service_name = None)))

	#Cenário 2:
	if(scenario == 2):

		#print(decrypt_message(b'gAAAAABfsspyDs_mXzXkYAkb4YfPgtf0P4-mTMNDiTZKBJwQPjIcNEzmiwK_-K502Y4LFRwMnR3HJ0R13TS0d-9GaTsW0sW9GmG729hueo88QUPaL7WXjo_mYh-HqpqiFLAJv12sWPgtL9puAhyIUd6Z4WQBkWgtIwDUNd-DtEsuqbHKQesqf1SKv2GY3butyXNZu_Dnh-C0',key_file))

		message = "'test message cenário 2'"
		generate_key(None, "test_cenario2")
		encrypt_message(message, None, None, "test_cenario2")
		print(str(decrypt_file(None, None, service_name = "test_cenario2")))

