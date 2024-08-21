import pandas as pd
import smtplib
import os
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

name_account = "Asesorias MJM"
email_account = "xxxxxxx@gmail.com"
password_account = "xxxx xxxx xxxx xxxx" #pass of aplication

server= smtplib.SMTP("smtp.gmail.com",587)
server.starttls()
server.login(email_account,password_account)

email_df = pd.read_excel("base_datos.xlsx") #database of mails (108 per run)
all_name = email_df['Empresa'] #database column
all_email = email_df['Correo']  #database column

for i in range(100):
    name = all_name[i]
    email = all_email[i]
    message=MIMEMultipart()
    message['subject'] = "Aqui va el Asunto"
    message['from'] = name_account
    message['To'] = email

    
    html = f"""
    <html>
    <body>
        <p><b>Estimado(a)</b> Ingeniero(a),</p>
        <p>Somos una .</p>

        <p><b>En , ofrecemos los siguientes servicios:</b></p>
         <img src="cid:image1" alt="imagename" style="width: 600px; max-width: 100%; height: auto;">
        <p>- Implementación y mejoramiento de <u>planes y programas de aseguramiento</u></p>
        <p>Nos encantaría coordinar una reunión para revisar sus necesidades específicas y evaluar el apoyo que podríamos brindarle.</p>

        <p>Cordial saludo,</p>
        <img src="cid:image2" alt="Firma">
    </body>
    </html>
    """
    body=MIMEText(html, 'html')
    message.attach(body)
    with open('imagename', 'rb') as image:
        image = MIMEImage(image.read())
        image.add_header('Content-ID', '<image1>')
        image.add_header('Content-Disposition', 'inline', filename='namefile.png')
        message.attach(image)


    with open('F.png', 'rb') as image:
        image2= MIMEImage(image.read())
        image2.add_header('Content-ID', '<image2>')
        image2.add_header('Content-Disposition', 'inline', filename='namefile.png')
        message.attach(image2)

    try:
        server.sendmail(email_account, [email], message.as_string())
    except Exception:
        print('Could not send email to {}. Error: {}\n'.format(email, str(Exception)))
    print(i,[email])
    time.sleep(30) #enviar cada 30 segundos
    
server.close()
