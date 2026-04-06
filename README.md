# Sistema Automatizado de Processamento e Análise de Pedidos

📌 Descrição

Este projeto consiste em um robô de automação (RPA) desenvolvido em Python que realiza, de forma totalmente automatizada, o processamento de pedidos a partir de planilhas armazenadas no Google Drive.

O sistema executa um fluxo completo que envolve:

* Acesso ao Google Drive
* Download de planilhas
* Processamento e análise de dados
* Geração de relatórios em PDF
* Envio automático por e-mail
  
# Fluxo do Sistema

O robô segue o seguinte fluxo:

* Acessa o Google Drive via navegador
* Realiza login na conta
* Localiza e baixa as planilhas:
  - Clientes
  - Pedidos
  - Produtos
* Processa os dados utilizando Python (Pandas)
* Gera relatórios com gráficos (Matplotlib)
* Exporta o relatório em PDF
* Envia o relatório por e-mail automaticamente
  
# Arquitetura

O sistema é dividido em três módulos principais:

* Coleta de Dados (RPA - Playwright)
* Processamento e Análise (Pandas + Matplotlib)
* Entrega (SMTP/Gmail)

# Estrutura do Projeto
```
📁 projeto/
│
├── desafio_py.py           # Acesso ao Drive e download das planilhas
├── relatorio.py            # Processamento e geração do PDF
├── enviar_relatorio.py     # Envio do e-mail com anexo
├── main.py                 # Orquestrador do sistema
├── requirements.txt        # Dependências do projeto
├── .env                    # Credenciais (não versionar)
│
├── Cliente.xlsx
├── Pedidos.xlsx
├── Produtos.xlsx
│
└── relatorio_final.pdf
```

# Pré-requisitos

Antes de executar o projeto, você precisa ter:

* Python 3.10 ou superior
* Google Chrome instalado
* Conta Google válida
* Acesso ao Google Drive com as planilhas
* Conta Gmail com acesso SMTP habilitado
  

# Instalação
1. Clone o repositório git clone [https://github.com/fanitsouza/Robo_Automacao_AnalisePedidos.git]
1. Crie o ambiente virtual python -m venv .venv
2. Ative o ambiente virtual .venv\Scripts\activate

* Caso dê erro de permissão: Set-ExecutionPolicy -Scope Process-ExecutionPolicy Bypass
1. Instale as dependências pip install -r requirements.txt
1. Instale os navegadores do Playwright playwright install
   

# Configuração de Credenciais

* Crie um arquivo .env na raiz do projeto:
  - GMAIL_USER=seu_email@gmail.com
  - GMAIL_PWD=sua_senha_de_app

# Importante:

* Não use sua senha normal do Gmail
* Gere uma Senha de App no Google
  
# Estrutura das Planilhas

📁 Cliente.xlsx

| id_cliente | nome | email | numero |

📁 Produtos.xlsx

| id_produto | descricao | valor |

📁 Pedidos.xlsx

| id_pedido | data | id_cliente | id_produto | quantidade | valor_total | status |

# Como Executar

Execute o arquivo principal:

### python main.py


O robô irá:

✔ Baixar as planilhas

✔ Processar os dados

✔ Gerar gráficos e análises

✔ Criar um arquivo PDF

✔ Enviar o relatório por e-mail


# Tecnologias Utilizadas
Python
Playwright → automação web
Pandas → manipulação de dados
Matplotlib → geração de gráficos
ReportLab / FPDF → geração de PDF
SMTP (Gmail) → envio de e-mails
Dotenv → gerenciamento de credenciais

# Limitações
Dependência de layout do Google Drive (pode mudar)
Necessidade de desativar verificação em duas etapas ou usar App Password
Execução depende de internet estável
Automação sensível a mudanças de interface

# Melhorias Futuras
Dashboard interativo (Power BI ou Streamlit)
Integração com banco de dados
Monitoramento automático de e-mails
Logs estruturados (arquivo .log)
Tratamento avançado de exceções
Execução agendada (Task Scheduler / Cron)

# Autora

Fani Tamires de Souza Batista

# Licença

Este projeto é de uso acadêmico e educacional.