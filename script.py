import pandas as pd
import datetime as dt
import smtplib
import os
from dotenv import load_dotenv  
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

load_dotenv()

# Lendo o arquivo Excel
# Quem ainda não pagou e quem está com data vencida
tabela = pd.read_excel('Contas a Receber.xlsx')

diaAtual = dt.datetime.now()

tabela_devedores = tabela.loc[tabela['Status']=='Em aberto']
tabela_devedores = tabela_devedores.loc[tabela_devedores['Data Prevista para pagamento']<diaAtual]
# print(tabela_devedores)

dados = tabela_devedores[['Valor em aberto', 'Data Prevista para pagamento', 'E-mail', 'NF']].values.tolist()

for dado in dados:
    sender_email = os.environ['EMAIL'] 
    receiver_email = dado[2]
    nf = dado[3]
    prazo = dado[1]
    prazo = prazo.strftime("%d/%m/%Y")
    valor = dado[0]
    password = os.environ['SENHA']
 
    subject = 'Atraso no Pagamento'
    body = f'''
    Prezado Cliente,
    
    Verificamos um atraso no pagamento referente a NF {nf} com vencimento em {prazo} e valor total de R${valor:.2f}.
    Gostaríamos de verificar se há algum problema que necessite de auxílio da nossa equipe.

    Att,
    HSLplace

    '''

  # Configurar o email
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject
    message.attach(MIMEText(body.encode('utf-8'), 'plain', 'utf-8'))

    # Enviar email usando SMTP
    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())

    print("Email Sent!")





