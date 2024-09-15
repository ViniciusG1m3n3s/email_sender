import pandas as pd
import sendgrid
from sendgrid.helpers.mail import Mail
import re
import logging
from flask import Flask, request, render_template

app = Flask(__name__)

# Configuração do arquivo de log
logging.basicConfig(filename='envio_emails.log', level=logging.INFO)

# Função para validar o formato do e-mail
def validar_email(email):
    return re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', email) is not None  # Atualização aqui

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        df = pd.read_excel(file)
        print("Colunas disponíveis no DataFrame:", df.columns)
        print("Primeiras linhas do DataFrame:")
        print(df.head())

        # Remover e-mails duplicados
        df = df.drop_duplicates(subset='Email')

        # Filtrar apenas os e-mails válidos
        df = df[df['Email'].apply(validar_email)]  # Verifique se a coluna 'Email' está correta (sensível a maiúsculas)

        # Inicializar a API do SendGrid com sua chave de API
        sg = sendgrid.SendGridAPIClient(api_key='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  ')

        # Loop por cada linha da planilha
        for index, row in df.iterrows():
            nome = row['Nome']
            email = row['Email']
            processo = row['Processo']

            # Criar o conteúdo do e-mail personalizado
            conteudo = f"""
            Olá {nome},

            Este é um aviso referente ao processo número {processo}.
            Por favor, entre em contato se tiver dúvidas.

            Atenciosamente,
            Sua Empresa
            """

            # Criar o e-mail
            message = Mail(
                from_email='viniciusnintendista0103@gmail.com',
                to_emails=email,
                subject=f'Atualização do processo {processo}',
                html_content=conteudo
            )

            # Enviar o e-mail
            try:
                response = sg.send(message)
                logging.info(f'E-mail enviado para {email}')
            except Exception as e:
                logging.error(f'Erro ao enviar e-mail para {email}: {e}')

        return render_template('sucesso.html')

if __name__ == '__main__':
    app.run(debug=True)
