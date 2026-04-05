''' O arquivo baixar_planilhas contém os dados de acesso ao e-mail que armazena em seu drive as
planilhas que irão gerar o relatório. Utilizando o playwright o robô loga na conta com o e-mail e senha,
acessa o drive e faz o download das planilhas definidas na função.
'''

# Imports necessários para execução
from playwright.sync_api import sync_playwright
import time
import re
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from dotenv import load_dotenv
import os

load_dotenv()

# Dados de acesso ao drive com os arquivos
EMAIL = os.getenv("EMAIL_USER")
SENHA = os.getenv("EMAIL_PWD")

# Após o acesso ao drive, a função baixar_planilhas inicia o processo de download
def baixar_planilhas():
    with sync_playwright() as p:
        # Abre o navegador
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        print("Acessando página de login...")
        page.goto("https://accounts.google.com/")

        # Faz o login
        # Insere o e-mail
        page.wait_for_selector('input[type="email"]')
        page.fill('input[type="email"]', EMAIL)
        page.click('button:has-text("Avançar")')

        page.wait_for_timeout(3000)

        # Insere a senha
        page.wait_for_selector('input[type="password"]')
        page.fill('input[type="password"]', SENHA)
        page.click('button:has-text("Avançar")')

        print("Realizando login...")
        page.wait_for_timeout(5000)

        # Acessa o Google Drive
        print("Acessando Google Drive...")
        page.goto("https://drive.google.com/drive/my-drive")

        # Espera carregar a página completamente
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_selector("text=Cliente")

        print("Arquivos encontrados!")
        
        # Arquivos que deverão ser baixados
        arquivos_para_baixar = ["Cliente", "Pedidos", "Produtos"]

        # Função para clicar e baixar os arquivos
        for nome_arquivo in arquivos_para_baixar:
            print(f"Procurando e baixando: {nome_arquivo}...")
            
            try:
                # 1. Localiza o arquivo pelo nome exato e clica com o botão direito
                # Usa .first para pegar o primeiro que aparecer (caso haja duplicatas visuais)
                elemento_arquivo = page.get_by_text(nome_arquivo, exact=True).first
                elemento_arquivo.wait_for(state="visible", timeout=10000)
                elemento_arquivo.click(button='right')
                
                # 2. Aguardando o evento de download acontecer
                with page.expect_download(timeout=30000) as download_info:
                    # 3. Clica na opção do menu de contexto. 
                    # O Regex ignora se estiver escrito "Download", "download" ou "Fazer o download"
                    opcao_download = page.locator('div[role="menuitem"]:visible', has_text=re.compile("Baixar|Download", re.IGNORECASE)).first

                    opcao_download.click()

                # 4. Captura o download e salva na mesma pasta do script
                download = download_info.value
                nome_sugerido = download.suggested_filename
                download.save_as(f"./{nome_sugerido}")
                
                print(f"Sucesso! Arquivo salvo como: {nome_sugerido}")
                
                # Pausa entre as execuções para evitar erros
                page.wait_for_timeout(1500)

            #Exceção caso o download não aconteça
            except Exception as e:
                print(f"Erro ao tentar baixar o arquivo '{nome_arquivo}': {e}")

        print("Todos os downloads foram processados!")
        
        time.sleep(3)
        #Fecha o browser
        browser.close()


if __name__ == "__main__":
    baixar_planilhas()