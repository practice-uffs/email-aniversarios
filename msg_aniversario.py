import pandas as pd
from email.message import EmailMessage
from email.utils import make_msgid
import smtplib

from dotenv import dotenv_values
config = dotenv_values(".env")

aniversariantes = 0

# trocar SHEET_ID para o arquivo compartilhado no GoogleDocs
url = f"https://docs.google.com/spreadsheets/d/{config['DATASHEET_ID']}/gviz/tq?tqx=out:csv&sheet={config['DATASHEET_NAME']}"

# dados vem das variáveis de ambiente
sender_name = config['MAIL_NAME']
sender_email = config['MAIL_USER']
sender_password = config['MAIL_PASS']
mail_from = sender_name + ' <' + sender_email + '>'

# servidor, porta e tipo de autenticação
server = smtplib.SMTP(config['MAIL_SERVER'], config['MAIL_PORT'])
server.starttls

# carrega data de hoje para verificar aniversariantes do dia
dia_de_hoje = pd.to_datetime('today').date().strftime('%d/%m')

# inicio do processamento da lista de servidores do campus Chapecó
# formato:
# Ativo	Nascimento	Prefixo	nome curto	nome completo	email
df = pd.read_csv(url)
df['Nascimento'] = pd.to_datetime(df['Nascimento'], dayfirst=True).dt.strftime('%d/%m')

attachment = 'niver.png'
with open(attachment, 'rb') as fd:
    img_data = fd.read()

# realiza uma query no dataframe em busca dos aniversariantes do dia
aniversariantes = df.query(f"(Ativo == 'sim') & (Nascimento == '{dia_de_hoje}')")

# se houverem aniversariantes envia a mensagem, senão imprime mensagem informativa
if len(aniversariantes) > 0:
    server.login(sender_email, sender_password)
    for i, row in aniversariantes.iterrows():
        print(row)

        message = EmailMessage()
        if row['Prefixo'] == 'nao':
            msg = row['Nome']
        else:
            msg = f"{row['Prefixo']}  {row['Nome']}"
        message['Subject'] = 'Feliz Aniversário'
        message['From'] = mail_from
        message['To'] = row['email']
        asparagus_cid = make_msgid()
        message.set_content(row['Nome_Completo'])

        message.add_alternative(f"""
        <html>
          <head></head>
          <body>
            <h4>Feliz Aniversário
            {str(msg)}
            !!!</h4>
            <img src="cid:{asparagus_cid[1:-1]}" />
          </body>
        </html>
        """, subtype='html')
        message.get_payload()[1].add_related(img_data, 'image', 'png', cid=asparagus_cid)


        server.send_message(message)
        print(msg + ': ' + row['email'])
    server.quit()
else:
    print(' Sem aniversários hoje!')