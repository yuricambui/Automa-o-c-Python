import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#importar base de dados

tabela_vendas = pd.read_excel("Vendas.xlsx")

#visualizar base de dados

pd.set_option("display.max_columns", None)
print(tabela_vendas)

print("-" * 50)
#faturamento por loja

faturamento_por_loja = tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()
print(faturamento_por_loja)

print("-" * 50)
#quantidade de produtos vendidos por loja

qtd_prod_vend_por_loja = tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja",).sum()
print(qtd_prod_vend_por_loja)

print("-" * 50)
#ticket médio por produto em cada loja

ticket_medio = (faturamento_por_loja["Valor Final"] / qtd_prod_vend_por_loja["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0: "Ticket Médio"})
print(ticket_medio)

#enviar um email com relatório

server_smtp = "smtp.gmail.com"
port = 587
sender_email = "E-mail do remetente"
password = "senha"

receive_email = "E-mail do destinatário"
subject = "E-mail automático em Python"
body = f'''
<p>Prezados, </p>

<p>Segue o Relatório de Vendas por Loja</p>

<p>Faturamento:</p>
{faturamento_por_loja.to_html(formatters={"Valor Final": "R${:,.2f}".format})}

<p>Quantidade Vendida</p>
{qtd_prod_vend_por_loja.to_html()}

<p>Ticket Médio dos Produtos em cada Loja</p>
{ticket_medio.to_html(formatters={"Ticket Médio": "R${:,.2f}".format})}

<p>Qualquer dúvida estou a disposição</p>

<p>Att.,</p>
<p>Yuri</p>
'''

message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receive_email
message["Subject"] = subject
message.attach(MIMEText(body, "html"))

try:
    server = smtplib.SMTP(server_smtp, port)
    server.starttls()

    server.login(sender_email, password)

    server.sendmail(sender_email, receive_email, message.as_string())
    print("E-mail enviado com sucesso")
except Exception as e:
    print(f"Houve algum erro: {e}")
finally:
    server.quit()


