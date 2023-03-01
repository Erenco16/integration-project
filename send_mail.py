import smtplib
from email.message import EmailMessage
import os
from dotenv import load_dotenv

load_dotenv()

def send_mail_with_excel(subject, content):

    sender_email = os.getenv("gmail_sender_email")
    recipient_email = os.getenv("gmail_receiver_email")
    app_password = os.getenv("gmail_app_password")
    excel_file = "dionaks_new_prices_and_stocks.xls"

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg.set_content(content)

    with open(excel_file, 'rb') as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype="xls", filename=excel_file)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(sender_email, app_password)
        smtp.send_message(msg)
