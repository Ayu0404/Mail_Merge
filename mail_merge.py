import os
from docx import Document
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv, dotenv_values

def get_message(template):
    msg=Document(template)
    return "\n".join([para.text for para in msg.paragraphs])


def get_email_data(email_list):
    return pd.read_excel(email_list)


def setup_server():
    server=smtplib.SMTP(os.getenv('smtp_server'),os.getenv('smtp_port'))
    server.starttls()
    server.login(os.getenv('username'),os.getenv('pswd'))
    return server


load_dotenv()

def mail_merge(template, email_list, subject):
    message=get_message(template)
    dataframe=get_email_data(email_list)
    server=setup_server()

    for idx,row in dataframe.iterrows():
        f_name=row.iloc[0]
        email_id=row.iloc[1]
        
        email=MIMEMultipart()
        email['from']=f'Ayushi Rai <{os.getenv('username')}>'
        email['to']=email_id
        email['subject']=subject
        greeting=message.replace('{{name}}',f_name)
        email.attach(MIMEText(greeting,'plain'))

        server.send_message(email)
    
    server.quit()
    print('Done!!')


template='template.docx'
email_list='list.xlsx'
subject='Subject 123'

mail_merge(template,email_list,subject)
