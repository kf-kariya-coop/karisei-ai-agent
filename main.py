import imaplib
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import decode_header
import re
import time
import os
from openai import OpenAI

GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")

client = OpenAI(api_key=OPENAI_API_KEY)

SYSTEM_PROMPT = """あなたはかりや生協（正式名称：かりや愛知中央生活協同組合）のAIスタッフです。
メールで届いた依頼に対して、組合職員として丁寧に対応してください。

【基本情報】
- 組織名：かりや生協
- あなたのメールアドレス：karisei.ai.agent@gmail.com
- あなたの名前：AIスタッフ

【対応できること】
- 質問への回答・調査
- 文書・資料の作成
- 日程調整（関係者へのメール文面を作成してお送りします）
- その他業務上の依頼

【日程調整について】
日程調整の依頼があった場合は、関係者に空き時間を確認するためのメール文面を作成し、
送信先と内容を明記して返信してください。

必ず丁寧な日本語で返信してください。署名は「かりや生協 AIスタッフ」としてください。"""


def decode_str(s):
    decoded = decode_header(s)
    result = ""
    for part, charset in decoded:
        if isinstance(part, bytes):
            charset = charset or "utf-8"
            try:
                result += part.decode(charset)
            except Exception:
                result += part.decode("utf-8", errors="replace")
        else:
            result += part
    return result


def get_email_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    body = part.get_payload(decode=True).decode("utf-8", errors="replace")
                    break
                except Exception:
                    pass
    else:
        try:
            body = msg.get_payload(decode=True).decode("utf-8", errors="replace")
        except Exception:
            pass
    return body


def generate_reply(sender_name, subject, body):
    user_message = f"""以下のメールが届きました。返信を作成してください。

送信者：{sender_name}
件名：{subject}

本文：
{body}
"""
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_message}
        ]
    )
    return response.choices[0].message.content


def send_reply(to_address, subject, body):
    msg = MIMEMultipart()
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = to_address
    if subject.startswith("Re:"):
        msg["Subject"] = subject
    else:
        msg["Subject"] = f"Re: {subject}"
    msg.attach(MIMEText(body, "plain", "utf-8"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, to_address, msg.as_string())
    print(f"返信送信完了: {to_address}")


def check_and_reply():
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        mail.select("inbox")

        _, message_ids = mail.search(None, "UNSEEN")
        ids = message_ids[0].split()

        if not ids:
            print("未読メールなし")
            mail.logout()
            return

        print(f"未読メール {len(ids)} 件")

        for msg_id in ids:
            _, msg_data = mail.fetch(msg_id, "(RFC822)")
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            sender = decode_str(msg.get("From", ""))
            subject = decode_str(msg.get("Subject", "（件名なし）"))
            body = get_email_body(msg)

            email_match = re.search(r'<(.+?)>', sender)
            sender_email = email_match.group(1) if email_match else sender
            sender_name = sender.split("<")[0].strip() if "<" in sender else sender

            print(f"処理中: {sender_email} / {subject}")

            if sender_email == GMAIL_ADDRESS:
                print("自分自身からのメールのためスキップ")
                continue

            reply_body = generate_reply(sender_name, subject, body)
            send_reply(sender_email, subject, reply_body)

        mail.logout()

    except Exception as e:
        print(f"エラー発生: {e}")


def main():
    print("かりや生協 AIエージェント 起動")
    print(f"メールアカウント: {GMAIL_ADDRESS}")

    while True:
        print("メール確認中...")
        check_and_reply()
        print("5分後に再確認します\n")
        time.sleep(300)


if __name__ == "__main__":
    main()
