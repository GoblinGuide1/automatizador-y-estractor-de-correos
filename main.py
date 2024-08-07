import smtplib
import os
import pandas as pd
from dotenv import load_dotenv
from email.message import EmailMessage
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#
#exel = pd.read_excel("C:/Users/Jorge/Desktop/python/AutomaticEmail/CorreosPractica.xlsx", header = 0)
#df = exel.loc[:, exel.columns == 'Correos']
#

correo = 'C:/Users/Jorge/Desktop/python/AutomaticEmail/correo.txt'
load_dotenv()
password = os.getenv("password") # vareable que guarda la contraseña dad por google
#datos para enviar el correo
email_sender = "Jorgebiuza@gmail.com" # correo emitente



# le indicamos a pandas que archivo va leer por medio de la ruta
exel = pd.read_excel("C:/Users/Jorge/Desktop/python/AutomaticEmail/CorreosPrueba.xlsx", header = 0)

# Cargamos el archivo Excel en un DataFrame
try:
    # Imprimimos la columna "Correos" si existe
    if 'Correos' in exel.columns:
        correos = exel['Correos']
        print(correos)
    else:
        print("La columna 'Correos' no se encuentra en el archivo Excel.")
except FileNotFoundError:
    print(f"El archivo en la ruta {exel} no fue encontrado.")
except Exception as e:
    print(f"Ocurrió un error al leer el archivo Excel: {e}")


email_reciver = correos


subject = "Solicitud de Práctica profesional"
archivoAdjuntado = "C:/Users/Jorge/Desktop/python/AutomaticEmail/Jorge Biuza Brenes CV.pdf"
nombre_adjunto = "Jorge Biuza Brenes CV.pdf"
with open(correo, 'r', encoding='utf-8') as fp:
    msg = MIMEMultipart()
   # msg.set_content(fp.read())
    text = MIMEText(fp.read(), 'plain', 'utf-8')
    # Adjunta la parte de texto al mensaje multipart
    msg.attach(text)

# Abrimos el archivo que vamos a adjuntar
archivo_adjunto = open(archivoAdjuntado, 'rb')
# Creamos un objeto MIME base
adjunto_MIME = MIMEBase('application', 'octet-stream')
# Y le cargamos el archivo adjunto
adjunto_MIME.set_payload((archivo_adjunto).read())
# Codificamos el objeto en BASE64
encoders.encode_base64(adjunto_MIME)
# Agregamos una cabecera al objeto
adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)

 

# em = EmailMessage()
msg["From"] = email_sender
msg["To"] = ", ".join(email_reciver)  # Convert the list to a comma-separated string
msg["subject"] = "Solicitud de Práctica profesional"
# Y finalmente lo agregamos al mensaje
msg.attach(adjunto_MIME)

# Envía el mensaje a través del servidor SMTP de Gmail.

try:
    
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()  # Inicia la conexión TLS
    s.login(email_sender, password)  # Inicia sesión en el servidor SMTP
    s.send_message(msg)
    s.quit()
    print("Correo enviado exitosamente")
except Exception as e:
    print(f"Ocurrió un error: {e}")

