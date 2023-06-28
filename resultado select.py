import pyodbc
import smtplib
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configurações de conexão com o banco de dados
server = 'IP BANCO DE DADOS'
database = 'BANCO DE DADOS'
username = 'USUARIO'
password = 'SENHA'
consulta_sql = '''COMANDO BANO DE DADOS, SOMENTE SELECT'''

# Configurações de conexão com o servidor de e-mail
smtp_host = 'smtp.dominio.com.br'
smtp_port = 'portasmtp'
smtp_username = 'email@email.com.br'
smtp_password = 'senha'

# Configurações dos e-mails
remetente = 'email@email.com.br'
destinatarios = ['email@email.com.br','email@email.com.br']
assunto = 'Assunto do email'

# Conectando-se ao banco de dados e executando a consulta
conexao = pyodbc.connect(f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}')
cursor = conexao.cursor()
cursor.execute(consulta_sql)
resultados = cursor.fetchall()
conexao.close()

# Obter a data e hora atual
data_hora_atual = datetime.datetime.now()

# Converter dia,mês,ano hora,minuto,segundo
data_hora_formatada = data_hora_atual.strftime('%d/%m/%Y %H:%M:%S')

# Preparando o conteúdo do e-mail
mensagem = MIMEMultipart()
mensagem['From'] = remetente
mensagem['To'] = ', '.join(destinatarios)
mensagem['Subject'] = assunto

# Formatando os resultados da consulta como uma tabela
corpo_email = f'<h2>Resultado dos usuários sem código de RH de acordo com o dia: {data_hora_formatada}</h2>'
corpo_email += '<table style="border-collapse: collapse; padding: 5px;">'
corpo_email += '<tr>'
# Para acrescentar mais colunas de acordo com seu coomando, adicione outra linha igual a baixo, ela será o nome da coluna
corpo_email += '<th style="border: 1px solid black; padding: 5px;">Código Senior</th>'
corpo_email += '</tr>'
for resultado in resultados:
    corpo_email += '<tr>'
    # Caso acrescente mais coluna acima, acrescentar outra linha abaixo porém com resultado[+1 do anterior] se for 3 colunas, vai ter que ter o resultado 0, 1, 2
    corpo_email += f'<td style="border: 1px solid black; padding: 5px;">{resultado[0]}</td>'
    corpo_email += '</tr>'
corpo_email += '</table>'
corpo_email += '<br>============================== E-mail automático para verificação de código de RH =============================='

# Adicionando o corpo do e-mail à mensagem
mensagem.attach(MIMEText(corpo_email, 'html'))

# Enviando o e-mail
with smtplib.SMTP_SSL(smtp_host, smtp_port) as smtp:
    smtp.login(smtp_username, smtp_password)
    smtp.send_message(mensagem)