import smtplib
import config
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_email(receiver, anexo, subject = "Você recebeu um certificado do CAAda!", message = "Segue em anexo seu certificado"):
    try:
        server  = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(config.EMAIL, config.PASSWORD)
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = config.EMAIL
        msg['To'] = receiver
        msg.attach(MIMEText(message, "plain"))
        with open(anexo, "rb") as f:
            attach = MIMEApplication(f.read(), _subtype='pdf')
        attach.add_header('Content-Disposition', 'attachment', filename = anexo)
        msg.attach(attach)
        server.send_message(msg)
    except:
        print("Email não enviado para " + receiver)
    else:
        print("Email enviado para " + receiver + " com sucesso!")
    finally:
        server.quit()

