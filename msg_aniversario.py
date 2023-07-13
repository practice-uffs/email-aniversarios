import numpy as np
import pandas as pd
from envelope import Envelope
from dotenv import dotenv_values

config = dotenv_values(".env")

images = [f"card{i+1}.jpeg" for i in range(9)]

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

# realiza uma query no dataframe em busca dos aniversariantes do dia
aniversariantes = df.query(f"(Ativo == 'sim') & (Nascimento == '{dia_de_hoje}')")



# se houverem aniversariantes envia a mensagem, senão imprime mensagem informativa
if len(aniversariantes) > 0:
    # servidor, porta e tipo de autenticação
    # para cada linha contendo um aniversariante
    for i, row in aniversariantes.iterrows():

        print(row)
        card = np.random.randint(1, 9, 1)[0]
        card_selecionado = images[card - 1]

        if row['Prefixo'] == 'nao':
            msg = row['Nome'].title()
        else:
            msg = f"{row['Prefixo']}  {row['Nome']}".title()

        # monta a mensagem
        html_message = f"""
                <html>
                <head></head>
                <body>
                    <h4>Feliz Aniversário {str(msg)}!</h4>
                    <img src="cid:{card_selecionado}" />
                </body>
                </html>
                """
        # TODO: ver uma forma mais elegante de iniciar o server e não duplicar os endereços
        server = Envelope(from_=f"{sender_name} <{sender_email}>").smtp(
            host=server_host,
            port=server_port,
            user=sender_email,
            password=sender_password,
            security='starttls'
        )

        server.attach(
            path=f"assets/{card_selecionado}", inline=True
        ).subject("Feliz Aniversário!").message(html_message).to(row['email']).send()

        server.smtp_quit()

        # informa sobre os emails enviados
        print(f"{msg}: {row['email']}")


else:
    print(' Sem aniversários hoje!')