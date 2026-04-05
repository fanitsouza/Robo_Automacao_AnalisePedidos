''' O arquivo enviar_relatorio.py é responsável por pegar o relatório com os dados gerados e enviar
para o responsável. Nesse código foi utilizado o smtplib para fazer a conexão no e-mail, e o
email.mime foi utilizado para fazer o envio do e-mail.
'''

# Imports necessários para execução
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

# Carrega as variáveis de ambiente
load_dotenv()

# Configuração do e-mail, senha, destinatário, assunto e o relatório
GMAIL_USER = os.getenv('GMAIL_USER')
GMAIL_PWD = os.getenv('GMAIL_PWD')

DESTINATARIO = "2026500347@ifam.edu.br"
ASSUNTO = "Relatório Automatizado de Pedidos"

CAMINHO_PDF = "relatorio_final.pdf"

# Definição do corpo do e-mail - mensagem que aparecerá no e-mail
CORPO_HTML = """
<html>
  <body>
    <h2>Relatório de Pedidos</h2>
    <p>Olá,</p>
    <p>Segue em anexo o relatório automatizado de pedidos gerado pelo robô RPA.</p>
    <p>O relatório contém:</p>
    <ul>
      <li>Resumo de pedidos</li>
      <li>Faturamento total</li>
      <li>Análise por status</li>
      <li>Top clientes</li>
    </ul>
    <br>
    <p>Este envio foi realizado automaticamente.</p>
    <p><b>Atenciosamente,</b><br></p>
  </body>
</html>
"""

# Função para criar a mensagem
def criar_msg():
    msg = MIMEMultipart()
    msg['Subject'] = ASSUNTO
    msg['From'] = GMAIL_USER
    msg['To'] = DESTINATARIO

    msg.attach(MIMEText(CORPO_HTML, 'html', 'utf-8'))

    return msg

# Função para carregar o arquivo que será enviado, no caso o relatorio_final
def carregar_anexo(msg, path):
    if not os.path.exists(path):
        print(f"Arquivo não encontrado: {path}")
        return

    with open(path, 'rb') as f:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(f.read())

    encoders.encode_base64(parte)

    nome_arquivo = "relatorio_final.pdf"
    parte.add_header(
        'Content-Disposition',
        f'attachment; filename="{nome_arquivo}"'
    )

    msg.attach(parte)
    print(f"Anexo adicionado: {nome_arquivo}")

# Função para enviar o e-mail, preenche os dados e envia
def enviar_email(msg):
    print("Conectando ao Gmail...")

    with smtplib.SMTP('smtp.gmail.com', 587) as servidor:
        servidor.starttls()
        servidor.login(GMAIL_USER, GMAIL_PWD)

        servidor.sendmail(
            from_addr=GMAIL_USER,
            to_addrs=[DESTINATARIO],
            msg=msg.as_string()
        )

    print("Email enviado com sucesso!")

# Executando e enviando o e-mail com o dado
def main():
    # Caso o e-mail ou a senha sejam inconsistente
    if not GMAIL_USER or not GMAIL_PWD:
        print("Erro: Verifique o .env")
        return

    print("Montando email...")

    msg = criar_msg()

    # Anexa o PDF
    carregar_anexo(msg, CAMINHO_PDF)

    # Envia
    enviar_email(msg)

if __name__ == "__main__":
    main()
