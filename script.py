import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from email.mime.base import MIMEBase
from email import encoders
import os


class Email:
    def __init__(self):
        self.subject = ''
        self.content = []
        self.email_password = []
        self.emails = []

    def iniciar(self):
        self.login()
        self.contentEmail()
        self.listEmails()
        self.email_preparation()
        self.send_email_me()

    def login(self):
        conta = pd.read_excel('login/email_senha.xlsx')            
        for linha in conta['Conta']:
            self.email_password.append(linha)
        self.email = self.email_password[0]
        self.password = self.email_password[1]
    
    def contentEmail(self):
        with open(r'email/conteudo.txt', 'r', encoding = 'utf-8') as arquivo:
            for linha in arquivo.read().splitlines():
                self.content.append(linha)
        
        self.subject += self.content[0]
        self.content.pop(0)
        self.content = '\n'.join(self.content)

    def listEmails(self):
        emails = pd.read_excel('email/emails.xlsx')
        for email in emails['email']:
            self.emails.append(email)

    def email_preparation(self):
        self.msg = MIMEMultipart()
        self.msg['Subject'] = self.subject
        self.msg['From'] = self.email
        self.msg['To'] = self.email
        self.msg.attach(MIMEText(self.content))
        self.msg.as_string().encode('utf-8')

        for i in os.listdir('email/anexos'):
            with open(f'email/anexos/{i}','rb') as file:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload((file).read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', "file; filename= %s" % i)
                self.msg.attach(part)

    def send_email_me(self):
        self.send()
        print(f'Um modelo do e-mail foi enviado para: {self.email}')
        self.quest = input('Enviar o e-mail para a lista de e-mails? [s/n]').strip().upper()
        if self.quest == 'SIM' or self.quest in 'S':
            self.send_emails()

    def send_emails(self):
        for email in self.emails:
            self.msg['To'] = str(email)
            self.send()
            print(f'Um e-mail foi enviado para: {email}')

    def send(self):
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(self.email, self.password)
            smtp.send_message(self.msg)


enviar_emails = Email()
enviar_emails.iniciar()
