import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate
import os

# 設定の読み込み
def load_email_settings():
    with open('mail/mail_address.json', 'r', encoding='utf-8') as f:
        email_settings = json.load(f)["email_settings"]
    with open('mail/mail_subject.txt', 'r', encoding='utf-8') as f:
        mail_subject = f.read().strip()
    with open('mail/mail_body.txt', 'r', encoding='utf-8') as f:
        body_template = f.read().strip()
    return email_settings, mail_subject, body_template

# SMTPサーバーの設定を読み込み
with open('settings/smtp.json', 'r', encoding='utf-8') as f:
    smtp_settings = json.load(f)

def send_email_with_attachment(email_settings, subject, body_template, result_data, attachment_path):
    msg = MIMEMultipart()
    msg['From'] = f"{email_settings['FromName']} <{email_settings['From']}>"
    msg['To'] = ", ".join(email_settings['TO'])
    msg['Cc'] = ", ".join(email_settings['CC'])
    msg['Subject'] = subject
    msg['Date'] = formatdate(localtime=True)

    # メール本文のプレースホルダを実際のデータで置き換え
    body_filled = body_template.format(
        total_count=result_data['totalCount'],
        succeed_count=result_data['succeededCount'],
        added_count=result_data['addedCount'],
        error_count=result_data['failedCount']
    )
    msg.attach(MIMEText(body_filled, 'plain', 'utf-8'))

    # ログファイルをメールに添付
    with open(attachment_path, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        filename = os.path.basename(attachment_path)
        part.add_header('Content-Disposition', 'attachment', filename=filename)  # Correct formatting for filename
        msg.attach(part)

    # SMTPサーバーを介してメールを送信
    with smtplib.SMTP(smtp_settings['smtp_host'], smtp_settings['smtp_port']) as server:
        # ユーザー名とパスワードが設定されている場合のみ認証を行う
        smtp_user = smtp_settings.get('smtp_user')
        smtp_password = smtp_settings.get('smtp_password')
        if smtp_user and smtp_password:
            server.login(smtp_user, smtp_password)
        server.send_message(msg)
