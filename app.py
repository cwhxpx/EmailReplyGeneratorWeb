from flask import Flask, render_template, request, redirect, url_for, session
from openai import OpenAI
import imapclient
import email
from email.policy import default
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import re
import config

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # 用于Flask会话管理

openai_client = OpenAI(api_key=config.OPENAI_API_KEY)


def get_server_domain(email_account):
    domain = email_account.split('@')[1]
    if 'gmail' in domain:
        return 'imap.gmail.com', 'smtp.gmail.com'
    elif 'outlook' in domain or 'hotmail' in domain or 'live' in domain:
        return 'imap-mail.outlook.com', 'smtp.office365.com'
    elif 'yahoo' in domain:
        return 'imap.mail.yahoo.com', 'smtp.mail.yahoo.com'
    else:
        return None, None


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        email_account = request.form.get('email_account')
        email_pwd = request.form.get('email_pwd')
        imap_server, smtp_server = get_server_domain(email_account)

        if not imap_server or not smtp_server:
            return "Unsupported email domain", 400

        session['email_account'] = email_account
        session['email_pwd'] = email_pwd
        session['imap_server'] = imap_server
        session['smtp_server'] = smtp_server

        return redirect(url_for('emails'))
    return render_template('index.html')


@app.route('/emails')
def emails():
    emails = last_10_emails()
    return render_template('emails.html', emails=emails)


@app.route('/generate_reply', methods=['POST'])
def generate_reply():
    selected_subject = request.form.get('email_subject')
    msg_id = email_subjects_dict.get(selected_subject)

    if not msg_id:
        return "Email not found", 404

    email_message, email_body = get_email_content(msg_id)

    # Use ChatGPT API to generate the reply
    response = openai_client.chat.completions.create(
        model="gpt-4",
        messages=[
            {"role": "user", "content": "You are a professional email writer"},
            {"role": "assistant", "content": "Ok"},
            {"role": "user", "content": f"Create a reply to this email:\n{email_body}"}
        ]
    )

    reply_body = response.choices[0].message.content
    session['from_address'] = email_message['From']
    session['subject'] = email_message['Subject']

    return render_template('reply.html', reply_body=reply_body)


@app.route('/send_reply', methods=['POST'])
def send_reply():
    reply_body = request.form.get('reply_body')  # 获取用户修改后的回复内容
    from_address = session.get('from_address')
    subject = session.get('subject')
    email_account = session.get('email_account')
    email_pwd = session.get('email_pwd')
    smtp_server = session.get('smtp_server')

    # Create the email reply
    msg = MIMEMultipart()
    msg['From'] = email_account
    msg['To'] = from_address
    msg['Subject'] = f"Re: {subject}"
    msg.attach(MIMEText(reply_body, 'plain'))

    # Send the email using SMTP
    try:
        with smtplib.SMTP(smtp_server, 587) as server:
            server.starttls()
            server.login(email_account, email_pwd)
            server.send_message(msg)
            print("Reply sent successfully.")
            return redirect(url_for('emails'))
    except Exception as e:
        return f"Failed to send reply: {e}", 500


def last_10_emails():
    email_account = session.get('email_account')
    email_pwd = session.get('email_pwd')
    imap_server = session.get('imap_server')

    # 连接到IMAP服务器
    server = imapclient.IMAPClient(imap_server, ssl=True)
    server.login(email_account, email_pwd)
    # 选择邮箱文件夹
    server.select_folder('INBOX')
    # 搜索邮件（此处为搜索所有未删除的邮件）
    messages = server.search(['NOT', 'DELETED'])
    # 获取最新的10个邮件的UID
    latest_messages = messages[-10:] if len(messages) > 10 else messages
    # 获取这些邮件的内容和标志
    response = server.fetch(latest_messages, ['BODY[]', 'FLAGS'])
    emails = []
    global email_subjects_dict
    email_subjects_dict = {}  # Dictionary to store email subjects and their corresponding msg_ids

    # 处理并显示邮件
    for msgid, data in response.items():
        email_message = email.message_from_bytes(data[b'BODY[]'], policy=default)
        subject = email_message['subject']
        from_ = email_message['from']
        to = email_message['to']
        subject_info = f"Subject: {subject}, From: {from_}, To: {to}"
        emails.append(subject_info)
        email_subjects_dict[subject_info] = msgid  # Store the subject info and its corresponding msg_id

    # 断开连接
    server.logout()
    return emails


def get_email_content(msg_id):
    email_account = session.get('email_account')
    email_pwd = session.get('email_pwd')
    imap_server = session.get('imap_server')

    # Connect to IMAP server
    server = imapclient.IMAPClient(imap_server, ssl=True)
    server.login(email_account, email_pwd)
    server.select_folder('INBOX')

    msg_data = server.fetch(msg_id, ['RFC822'])
    raw_email = msg_data[msg_id][b'RFC822']
    email_message = email.message_from_bytes(raw_email, policy=default)
    email_body = ""
    if email_message.is_multipart():
        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                email_body = part.get_payload(decode=True).decode()
                break
    else:
        email_body = email_message.get_payload(decode=True).decode()

    server.logout()
    return email_message, email_body


if __name__ == '__main__':
    app.run(debug=True)
