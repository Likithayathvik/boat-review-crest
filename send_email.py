import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime

EMAIL_USERNAME = os.environ['EMAIL_USERNAME']
EMAIL_PASSWORD = os.environ['EMAIL_PASSWORD']
NOTIFY_EMAIL   = 'likitha.m@imaginemarketingindia.com'
XLSX_FILE      = 'Android_CrestApp_Review.xlsx'

def send():
    msg = MIMEMultipart()
    msg['From']    = EMAIL_USERNAME
    msg['To']      = NOTIFY_EMAIL
    msg['Subject'] = f"✅ boAt Crest App Weekly Review Report — {datetime.now().strftime('%d %b %Y')}"

    body = """Hello Sathya,

The weekly boAt Crest App review analysis is complete.

Please find the Excel report attached.

— boAt R&D QA Automation"""

    msg.attach(MIMEText(body, 'plain'))

    with open(XLSX_FILE, 'rb') as f:
        attachment = MIMEApplication(f.read(), _subtype='xlsx')
        attachment.add_header(
            'Content-Disposition', 'attachment',
            filename=f"Android_CrestApp_Review_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        )
        msg.attach(attachment)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(EMAIL_USERNAME, EMAIL_PASSWORD)
        server.send_message(msg)

    print(f"✅ Email sent to {NOTIFY_EMAIL}")

if __name__ == '__main__':
    send()
