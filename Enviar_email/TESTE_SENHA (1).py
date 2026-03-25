# pip install python-dotenv
# pip install pandas
# pip install MIMEText
# pip install smtplib
# Desenvolvido por Fábio A Cordeiro
# Em 28/02/2025

import os
from dotenv import load_dotenv #type:igore
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

load_dotenv()
BD_PASSWORD = os.getenv('BD_PASSWORD')
print(os.environ)
print(os.getenv('EMAIL_PASSWORD'))


#seu_email = "novosprojetosbr@gmail.com"
#sua_senha = "zcfbkzhxkdhvixcw"

