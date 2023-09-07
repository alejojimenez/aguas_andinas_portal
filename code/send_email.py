import os
import time
import datetime
import pandas as pd

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import base64
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pickle

from requests.auth import HTTPBasicAuth
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# Scopes para usar API de gmail
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly','https://www.googleapis.com/auth/gmail.send']

# Definir metodos para enviar correos
def create_message_with_attachment(sender, to, subject, message_text, files):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    msg = MIMEText(message_text, 'html')
    message.attach(msg)

    files = filename_formato_planilla
    part = MIMEBase('application', "xlsx")
    part.set_payload(open(files, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=files)
    message.attach(part)
    
    raw = base64.urlsafe_b64encode(message.as_bytes())
    raw = raw.decode()
    return {'raw': raw}

# Definir metodos para enviar correos no satisfactorios
def Create_Message_Without_Attachment(sender, to, subject, message_text):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    msg = MIMEText(message_text, 'html')
    message.attach(msg)

    raw = base64.urlsafe_b64encode(message.as_bytes())
    raw = raw.decode()
    return {'raw': raw}

def send_message(service, user_id, message):
    try:
        message = (service.users().messages().send(userId=user_id, body=message)
                   .execute())
        print('Correo enviado. ID: ' + message['id'])
        return message
    except errors.HttpError as error:
        print('Ha ocurrido un error: ' + str(error))

html_mail = ""
mail_sender = 'alexander@macrobots.com'

creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('../config/token.pickle'):
    with open('../config/token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('../config/credencial_client_secret.com.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('../config/token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('gmail', 'v1', credentials=creds)
#
##################################################

####### Proceso enviar correo electronico ########
#
filename_formato_planilla = "C:\\roda\\aguas_andinas_portal\\output\\Formato Planilla.xlsx"
#

html_mail = """<p>Estimados</p>

<p>Junto con saludar, informar que el bot mencionado en el asunto corrio satisfactoriamente, adjunto al presente archivo procesado.</p>

<p>Cualquier duda con respecto al archivo generado, comunicarse con el administrador del proceso.</p>

<p>Saludos</p>"""

print('Preparando envio de correo')
print('------------------------------------')
#
asunto = f'[bot] Reporte de Ejecución - Procesamiento Facturas/Boletas - Aguas Andinas - ' + str(datetime.datetime.today().strftime('%d-%m-%Y'))
#
# Lista de destinatarios - Leer el archivo de Excel
destination_1 = pd.read_excel('C:\\roda\\aguas_andinas_portal\\config\\destinatarios.xlsx')

# Extraer las direcciones de correo electrónico de la columna 'Destinatarios'
receptores_correo = destination_1['Destinatarios'].to_list()
print('Lista de destinatarios Arreglo: ', receptores_correo)

# Convertir la lista de direcciones en una cadena separada por comas
receptores_correo_str = ', '.join(receptores_correo)

# Receptor de pruebas
# receptores_correo_str = 'alexander@macrobots.com'
#
files = filename_formato_planilla
#
message = create_message_with_attachment(mail_sender, receptores_correo_str, asunto, html_mail, files)
send_message(service, mail_sender, message)
time.sleep(5)
# 
##################################################
