import pandas as pd
from envelope import Envelope

from dotenv import dotenv_values
config = dotenv_values(".env")

aniversariantes = 0

# trocar SHEET_ID para o arquivo compartilhado no GoogleDocs
url = f"https://docs.google.com/spreadsheets/d/{config['DATASHEET_ID']}/gviz/tq?tqx=out:csv&sheet={config['DATASHEET_NAME']}"

# dados vem das variáveis de ambiente
server_host = config['MAIL_SERVER']
server_port = config['MAIL_PORT']
sender_name = config['MAIL_NAME']
sender_email = config['MAIL_USER']
sender_password = config['MAIL_PASS']

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
    # servidor, porta e tipo de autenticação
    server = Envelope(
            from_=sender_email
        ).smtp(
            host=server_host,
            port=server_port,
            user=sender_email,
            password=sender_password,
            security='starttls')
    for i, row in aniversariantes.iterrows():
        print(row)

        if row['Prefixo'] == 'nao':
            msg = row['Nome']
        else:
            msg = f"{row['Prefixo']}  {row['Nome']}"

        html_message = f"""
                <html>
                <head></head>
                <body>
                    <h4>Feliz Aniversário {str(msg)}!</h4>
                    <img src="cid:niver.png" />
                </body>
                </html>
                """
        server.attach(
                path="niver.png",
                inline=True
            ).subject("Feliz Aniversário!").message(html_message).to(row['email']).send()

        print(msg + ': ' + row['email'])

    server.smtp_quit()

else:
    print(' Sem aniversários hoje!')