import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager  
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


#todo Configuração para usar o Webdriver

options = webdriver.ChromeOptions()
options.add_argument("--user-data-dir=C:/Users/eduar/AppData/Local/Google/Chrome/User Data/Default")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)  

#todo Deixa o navegador em tela cheia

driver.maximize_window()

#todo Abrir o Whatsapp 

driver.get("https://web.whatsapp.com")
print('Aguardando login no WhatsApp...')

#todo Espera até que o WhatsApp carregue a lista de conversas (indicando que o login foi feito)

WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, "//div[@role='textbox']")))

print('Login detectado! Continuando a automação...')

#todo Lendo a planilha de contatos

planilha = pd.read_excel("contatos.xlsx")
print(f"{len(planilha)} contatos encontrados!")

#todo Loop para enviar mensagens

for index, row in planilha.iterrows():
    nome = row["Nome"]
    telefone = str(row["Telefone"])  #todo String
    mensagem = f"Olá, {nome}! Esta é uma mensagem automática enviada via Selenium."

    try:
        #todo Clicar no botão Nova Conversa
        nova_conversa = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//span[@data-icon='new-chat-outline']"))
        )
        nova_conversa.click()

        #todo Após clicar, espera a barra de pesquisa
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
        )

        #todo Digitar o número do contato na busca da nova conversa
        
        search_box.send_keys(telefone)
        time.sleep(2)  #todo Tempo para o WhatsApp encontrar o contato
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)

        #todo Campo de mensagem
        message_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='10']"))
        )
        message_box.send_keys(mensagem)
        message_box.send_keys(Keys.ENTER)
        time.sleep(2)  #todo Pequeno delay para evitar bloqueio
        
        print(f"✅ Mensagem enviada para {nome} ({telefone})")

    except Exception as e:
        print(f"❌ Erro ao enviar para {nome} ({telefone}): {str(e)}")

#todo Fechar navegador após envio

print("Todas as mensagens foram enviadas!")
driver.quit()
