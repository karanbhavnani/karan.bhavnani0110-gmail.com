
import smtplib, yaml
from email.mime.text import MIMEText

def send_email_notification(subject, body):
    c=yaml.safe_load(open("config.yaml"))
    msg=MIMEText(body)
    msg['Subject']=subject; msg['From']=c['email_sender']; msg['To']=c['email_receiver']
    s=smtplib.SMTP(c.get('smtp_server','smtp.gmail.com'),c.get('smtp_port',587))
    s.starttls(); s.login(c['email_sender'],c['email_password'])
    s.send_message(msg); s.quit()
