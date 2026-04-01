import imaplib
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import decode_header
import re
import os
from datetime import date, timedelta
from openai import OpenAI
from supabase import create_client

GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

client = OpenAI(api_key=OPENAI_API_KEY)
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

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
- 職員情報の登録・更新・削除（免許更新期限、ドライブレコーダー装着状況など）
- その他業務上の依頼

【職員情報の管理について】
以下のような依頼はデータベースを操作してください：
- 「○○ ○○、職員コード：XXX、メール：xxx@karisei.jp、免許更新：2027年3月を登録して」→ 新規登録
- 「職員コードXXXの免許更新期限を2028年5月に変更して」→ 更新
- 「○○ ○○を退職につき削除して」→ 削除
操作結果を必ず返信で報告してください。

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


def get_db_context():
    """メール処理時にDBの操作が必要か判断するためのコンテキスト情報を返す"""
    return """
【データベース操作コマンド】
メール本文に以下のキーワードが含まれる場合、対応するDB操作を指示してください：
- 「登録」「追加」→ INSERT
- 「更新」「変更」「修正」→ UPDATE
- 「削除」「退職」→ DELETE
- 「一覧」「確認」「検索」→ SELECT
"""


def handle_db_operation(operation, data):
    """データベース操作を実行する"""
    try:
        if operation == "insert":
            result = supabase.table("staff_licenses").insert(data).execute()
            return f"登録完了：{data.get('name', '')}さんのデータを登録しました。"
        elif operation == "update":
            staff_code = data.pop("staff_code", None)
            if not staff_code:
                return "エラー：職員コードが指定されていません。"
            result = supabase.table("staff_licenses").update(data).eq("staff_code", staff_code).execute()
            return f"更新完了：職員コード {staff_code} のデータを更新しました。"
        elif operation == "delete":
            staff_code = data.get("staff_code")
            name = data.get("name")
            if staff_code:
                result = supabase.table("staff_licenses").delete().eq("staff_code", staff_code).execute()
                return f"削除完了：職員コード {staff_code} のデータを削除しました。"
            elif name:
                result = supabase.table("staff_licenses").delete().eq("name", name).execute()
                return f"削除完了：{name}さんのデータを削除しました。"
            else:
                return "エラー：削除対象の職員コードまたは氏名が指定されていません。"
    except Exception as e:
        return f"データベースエラー：{e}"


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


def send_license_reminders():
    """免許更新期限が3ヶ月以内の職員にリマインドメールを送信する"""
    try:
        today = date.today()
        three_months_later = today + timedelta(days=90)

        result = supabase.table("staff_licenses").select("*").execute()
        staff_list = result.data

        reminded_count = 0
        for staff in staff_list:
            if not staff.get("license_expiry_date") or not staff.get("email"):
                continue

            expiry_date = date.fromisoformat(staff["license_expiry_date"])

            if today <= expiry_date <= three_months_later:
                days_left = (expiry_date - today).days
                send_license_reminder_email(staff, expiry_date, days_left)

                supabase.table("staff_licenses").update(
                    {"last_reminded_at": today.isoformat()}
                ).eq("id", staff["id"]).execute()

                reminded_count += 1

        print(f"免許更新リマインド送信完了：{reminded_count}件")

    except Exception as e:
        print(f"免許更新チェックエラー：{e}")


def send_license_reminder_email(staff, expiry_date, days_left):
    """個別の免許更新リマインドメールを送信する"""
    name = staff.get("name", "")
    to_address = staff["email"]

    body = f"""{name} さん

お疲れさまです。かりや生協 AIスタッフです。

運転免許証の更新期限が近づいていますのでお知らせします。

■ 更新期限：{expiry_date.strftime('%Y年%m月%d日')}（あと{days_left}日）

更新が完了しましたら、新しい有効期限をご連絡ください。
このメールに返信する形で「免許を更新しました。新しい期限は○年○月です。」とお送りいただければ、
自動的に記録を更新します。

よろしくお願いいたします。

かりや生協 AIスタッフ"""

    msg = MIMEMultipart()
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = to_address
    msg["Subject"] = "【かりや生協】運転免許証の更新期限のお知らせ"
    msg.attach(MIMEText(body, "plain", "utf-8"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, to_address, msg.as_string())
    print(f"免許更新リマインド送信：{name}（{to_address}）")


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
    print("免許更新リマインドチェック中...")
    send_license_reminders()
    print("メール確認中...")
    check_and_reply()
    print("完了")


if __name__ == "__main__":
    main()
