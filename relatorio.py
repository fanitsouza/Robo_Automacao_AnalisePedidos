''' O arquivo relatorio.py é responsável por reunir e limpar os dados e gerar o relatório.
Utilizando as bibliotecas pandas, matplotlib para ler os dados e gerar os gráficos é possível 
obter o relatório em pdf.
'''

# Imports necessários para execução
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet


# Realiza a leitura dos dados

df_clientes = pd.read_excel("Cliente.xlsx")
df_pedidos = pd.read_excel("Pedidos.xlsx")
df_produtos = pd.read_excel("Produtos.xlsx")

print("Dados carregados com sucesso!")

# Realiza o tratamento dos dados das diferentes planilhas

df = df_pedidos.merge(df_clientes, on="id_cliente")
df = df.merge(df_produtos, on="id_produto")

print("Dados integrados!")

# Métricas que irão compor o relatório

total_pedidos = len(df)
faturamento_total = df["valor_total"].sum()

pedidos_por_status = df["status"].value_counts()

top_clientes = df.groupby("nome")["valor_total"].sum().sort_values(ascending=False).head(5)

# Gera gráficos utilizando o Matplotlib

# Gráfico 1: pedidos por status - Cada pedido possui um status e ele reuni esses dados
plt.figure()
pedidos_por_status.plot(kind="bar")
plt.title("Pedidos por Status")
plt.ylabel("Quantidade")
plt.tight_layout()
plt.savefig("grafico_status.png")
plt.close()

# Gráfico 2: top clientes - clientes que mais gastaram na loja
plt.figure()
top_clientes.plot(kind="bar")
plt.title("Top 5 Clientes por Valor")
plt.ylabel("Valor Total")
plt.tight_layout()
plt.savefig("grafico_clientes.png")
plt.close()

print("Gráficos gerados!")

# Gera o relatório em PDF com os dados gerados anteriormente

doc = SimpleDocTemplate("relatorio_final.pdf")
styles = getSampleStyleSheet()

conteudo = []

#Organização do PDF
# Título
conteudo.append(Paragraph("Relatório de Pedidos", styles["Title"]))
conteudo.append(Spacer(1, 10))

# Data
data_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
conteudo.append(Paragraph(f"Gerado em: {data_atual}", styles["Normal"]))
conteudo.append(Spacer(1, 10))

# Métricas
conteudo.append(Paragraph(f"Total de pedidos: {total_pedidos}", styles["Normal"]))
conteudo.append(Paragraph(f"Faturamento total: R$ {faturamento_total:.2f}", styles["Normal"]))
conteudo.append(Spacer(1, 10))

# Status
conteudo.append(Paragraph("Pedidos por status:", styles["Heading2"]))
for status, qtd in pedidos_por_status.items():
    conteudo.append(Paragraph(f"{status}: {qtd}", styles["Normal"]))

conteudo.append(Spacer(1, 20))

# Imagens
conteudo.append(Paragraph("Gráfico de Status:", styles["Heading2"]))
conteudo.append(Image("grafico_status.png", width=400, height=250))

conteudo.append(Spacer(1, 20))

conteudo.append(Paragraph("Top Clientes:", styles["Heading2"]))
conteudo.append(Image("grafico_clientes.png", width=400, height=250))

# PDF Gerado
doc.build(conteudo)

print("Relatório PDF gerado com sucesso!")