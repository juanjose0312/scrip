import smtplib
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from datetime import datetime, timedelta

class Envio_de_correo:
    def __init__(self, recipients, subject, body, files):
        self.sender = "salesland.mov.informes@gmail.com"
        self.recipients = recipients
        self.subject = subject
        self.body = body
        self.files = files

    def send_email(self):
        msg = MIMEMultipart()
        msg['From'] = self.sender
        msg['To'] = ', '.join(self.recipients)
        msg['Subject'] = self.subject

        fecha_actual = datetime.now() 
        fecha_ayer = fecha_actual - timedelta(days=1)
        fecha_ayer = fecha_ayer.strftime("%d/%m/%Y")

        body = f"""----HOLA EQUIPO----\n
        ¡Espero que se encuentre bien!\n
        A continuacion les comparto el consolidado de bases actualizado al {fecha_ayer}.\n
        Este archivo contiene información sobre las bases consolidadas.\n
        """
        msg.attach(MIMEText(body, 'plain'))

        for file in self.files:
            attachment = open(file, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {file}")
            msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(self.sender, "yyxi qrwk grnj bazc")

        text = msg.as_string()
        server.sendmail(self.sender, self.recipients, text)

        print("¡Correo enviado correctamente!")
        server.quit()



# Uso de la clase
#recipients = ['giovanniandresb@salesland.net','ivonperdomo@salesland.net','coordinadoroperativo@salesland.net','coordinadoradminffvv@salesland.net']
#recipients = ['cardonagiljuanjose@gmail.com']
#subject = "correo de prueba con el segundo metodo"
#fecha_actual = datetime.now() 
#fecha_ayer = fecha_actual - timedelta(days=1)
#fecha_ayer = fecha_ayer.strftime("%d/%m/%Y")
#body = f"""----HOLA EQUIPO----\n
#¡Espero que se encuentre bien!\n
#A continuacion les comparto el consolidado de bases actualizado al {fecha_ayer}.\n
#Este archivo contiene información sobre las bases consolidadas.\n
#"""
#files = ["hola.txt", "chao.txt"]
#email_sender = Envio_de_correo(recipients, subject, body, files)
#email_sender.send_email()
