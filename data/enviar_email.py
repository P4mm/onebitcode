import smtplib
import ssl
import mimetypes
from email.message import EmailMessage

# 1 - Dados do E-mail
password = open("senha", "r").read()
from_email = "pamelag.gabi@gmail.com"
to_email = "harveypamela@hotmail.com"
subject = "Relatório de Vendas"
body = "Este é o relatório de vendas de ontem."

# 2 - Criando o e-mail
message = EmailMessage()
message["From"] = from_email
message["To"] = to_email
message["Subject"] = subject

message.set_content(body)
safe_password = ssl.create_default_context()

# 3 - Anexando arquivo com nome personalizado
anexo = "C:/Users/User/Desktop/python_estudos/minicurso_Onebitecode/data/teste.xlsx"
nome_personalizado = "relatorio_vendas.xlsx"
mime_type, mime_subtype = mimetypes.guess_type(anexo)[0].split("/")
with open(anexo, 'rb') as a:
    message.add_attachment(a.read(), maintype=mime_type, subtype=mime_subtype, filename=nome_personalizado)

# 4 - Enviando e-mail
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=safe_password) as smtp:
    smtp.login(from_email, password)
    smtp.sendmail(from_email, to_email, message.as_string())
    print("E-mail enviado com sucesso!")
