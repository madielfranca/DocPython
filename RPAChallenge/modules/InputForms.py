import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

class Form:
    def __init__(self, url_formulario, caminho_planilha):
        self.caminho_planilha = caminho_planilha
        self.url_formulario = url_formulario

    def inputForms(self):
        # Lê a planilha
        df = pd.read_excel(self.caminho_planilha)

        # Configura o Selenium
        options = webdriver.ChromeOptions()
        driver = webdriver.Chrome(options=options)

        driver.get(self.url_formulario)
        time.sleep(3)

        # Maximiza a janela do navegador
        driver.maximize_window()

        # Aguarda até que a página carregue
        wait = WebDriverWait(driver, 10)

        # Função para encontrar o campo de entrada associado ao rótulo
        def encontrar_campo_por_label(label_text):
            try:
                # Tenta encontrar o input seguindo diretamente o label
                campo = driver.find_element(By.XPATH, f"//label[contains(text(), '{label_text}')]/following-sibling::input")
                return campo
            except:
                try:
                    # Encontra o input através do atributo 'for' do label
                    label = driver.find_element(By.XPATH, f"//label[contains(text(), '{label_text}')]")
                    campo_id = label.get_attribute('for')
                    if campo_id:
                        campo = driver.find_element(By.ID, campo_id)
                        return campo
                except Exception as e:
                    print(f"Erro ao encontrar o campo '{label_text}': {e}")
                    return None
                
        start_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button")))
        start_button.click()

        # Preencher o formulário
        for index, row in df.iterrows():
            for col in df.columns:
                label_text = col.strip()  # Remova espaços extras do texto do label
                campo_formulario = encontrar_campo_por_label(label_text)
                if campo_formulario:
                    campo_formulario.clear()
                    campo_formulario.send_keys(str(row[col]))
                else:
                    print(f"Erro ao preencher o campo '{label_text}': Campo não encontrado.")
            
            # Envia o formulário
            try:
                time.sleep(3)
                submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit']")))
                submit_button.click()
            except Exception as e:
                print(f"Erro ao enviar o formulário: {e}")
            
            # Aguarda a página recarregar e os campos mudarem de posição
            time.sleep(2)

        # Fechaa o navegador
        time.sleep(10)
        driver.quit()

