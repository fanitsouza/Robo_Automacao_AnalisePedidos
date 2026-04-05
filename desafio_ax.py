from playwright.sync_api import sync_playwright
import time
import re
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet

EMAIL = "2026500347@ifam.edu.br"
SENHA = "AXacademy@2026"

def baixar_planilhas():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        print("Acessando página de login...")
        page.goto("https://accounts.google.com/")

        # ─── LOGIN ─────────────────────────────
        page.wait_for_selector('input[type="email"]')
        page.fill('input[type="email"]', EMAIL)
        page.click('button:has-text("Avançar")')

        page.wait_for_timeout(3000)

        page.wait_for_selector('input[type="password"]')
        page.fill('input[type="password"]', SENHA)
        page.click('button:has-text("Avançar")')

        print("Realizando login...")
        page.wait_for_timeout(5000)

        # ─── ACESSA GOOGLE DRIVE ──────────────
        print("Acessando Google Drive...")
        page.goto("https://drive.google.com/drive/my-drive")

        # NÃO usar networkidle
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_selector("text=Cliente")

        print("Arquivos encontrados!")

        arquivos_para_baixar = ["Cliente", "Pedidos", "Produtos"]

        for nome_arquivo in arquivos_para_baixar:
            print(f"Procurando e baixando: {nome_arquivo}...")
            
            try:
                # 1. Localiza o arquivo pelo nome exato e clica com o botão direito
                # Usamos .first para pegar o primeiro que aparecer (caso haja duplicatas visuais)
                elemento_arquivo = page.get_by_text(nome_arquivo, exact=True).first
                elemento_arquivo.wait_for(state="visible", timeout=10000)
                elemento_arquivo.click(button='right')
                
                # 2. Abre o bloco que "escuta" o evento de download
                with page.expect_download(timeout=30000) as download_info:
                    # 3. Clica na opção do menu de contexto. 
                    # Usamos Regex para ignorar se está escrito "Download", "download" ou "Fazer o download"
                    opcao_download = page.locator('div[role="menuitem"]:visible', has_text=re.compile("Baixar|Download", re.IGNORECASE)).first

                    opcao_download.click()

                # 4. Captura o download e salva na mesma pasta do script
                download = download_info.value
                nome_sugerido = download.suggested_filename
                download.save_as(f"./{nome_sugerido}")
                
                print(f"Sucesso! Arquivo salvo como: {nome_sugerido}")
                
                # Pausa breve para a interface do Drive "respirar" antes de ir para o próximo arquivo
                page.wait_for_timeout(1500)

            except Exception as e:
                print(f"Erro ao tentar baixar o arquivo '{nome_arquivo}': {e}")

        print("Todos os downloads foram processados!")
        
        time.sleep(3)
        browser.close()


if __name__ == "__main__":
    baixar_planilhas()