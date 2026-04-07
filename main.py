import imaplib
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import decode_header
import re
import os
import csv
import io
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

WORK_EMAIL_DOMAIN = "kariya-coop.or.jp"

SYSTEM_PROMPT = """あなたはかりや生協（正式名称：かりや愛知中央生活協同組合）のAIスタッフです。
メールで届いた依頼に対して、組合職員として丁寧に対応してください。
必ず丁寧な日本語で返信してください。署名は「かりや生協 AIスタッフ」としてください。

【個人情報の取り扱いルール】
・職員の個人メールアドレスは、いかなる場合も他の人に教えてはいけません
・職員間の日程調整や業務連絡には、組合メールアドレス（@kariya-coop.or.jp）のみ使用してください
・「○○さんのメールアドレスを教えて」という依頼には応じないでください
・組合メールアドレスも、業務上必要な場合のみ共有してください

【組織情報】
・組織名：かりや生協（正式名称：かりや愛知中央生活協同組合）
・職員数：約800名
・内部の呼称：「組合内」「職員」（「社内」「社員」は使わない）"""


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


def send_email(to_address, subject, body):
    msg = MIMEMultipart()
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = to_address
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, to_address, msg.as_string())


def get_csv_attachment(msg):
    """メールからCSV添付ファイルを取得する"""
    for part in msg.walk():
        content_disposition = part.get("Content-Disposition", "")

        # 添付ファイルのみを対象（メール本文は除外）
        if "attachment" not in content_disposition:
            continue

        # ファイル名を取得（エンコードされている場合もデコード）
        raw_filename = part.get_filename()
        filename = decode_str(raw_filename) if raw_filename else ""
        print(f"添付ファイル検出：{filename}")

        # CSV以外は除外
        if filename and not filename.lower().endswith(".csv"):
            continue

        payload = part.get_payload(decode=True)
        if not payload:
            continue

        # 文字コードを自動判定
        for encoding in ["utf-8-sig", "shift-jis", "cp932", "utf-8"]:
            try:
                content = payload.decode(encoding)
                if content.strip():
                    return content
            except Exception:
                continue

    return None


def handle_staff_master_import(sender_email, csv_content):
    """職員マスタCSVをインポートする"""
    try:
        reader = csv.reader(io.StringIO(csv_content))
        rows = list(reader)

        if not rows:
            send_email(sender_email, "【かりや生協】職員マスタ更新 - エラー",
                "CSVファイルが空です。確認してください。\n\nかりや生協 AIスタッフ")
            return

        print(f"CSVの行数：{len(rows)}")

        # 1行目がヘッダーかデータか判定
        first_row = rows[0]
        start_index = 1 if any(h in first_row for h in ["所属", "社員CD", "氏名"]) else 0

        inserted = 0
        updated = 0
        errors = 0
        new_staff_codes = []

        for row in rows[start_index:]:
            if len(row) < 9:
                continue
            try:
                # 列マッピング：A=所属, B=組織, C=所属部, D=所属課, E=職場名, F=社員CD, G=氏名, H=フリガナ, I=役職名
                staff_code_raw = row[5].strip()
                if not staff_code_raw or not staff_code_raw.isdigit():
                    continue

                staff_code = int(staff_code_raw)
                new_staff_codes.append(staff_code)

                data = {
                    "staff_code": staff_code,
                    "workplace_code": row[1].strip() if row[1].strip().isdigit() else None,
                    "department": row[2].strip(),
                    "section": row[3].strip(),
                    "workplace_name": row[4].strip(),
                    "name": row[6].strip(),
                    "name_kana": row[7].strip(),
                    "position_title": row[8].strip(),
                    "is_active": True,
                    "updated_at": date.today().isoformat()
                }

                # workplace_codeを整数に変換
                if data["workplace_code"]:
                    data["workplace_code"] = int(data["workplace_code"])

                existing = supabase.table("staff_master").select("staff_code").eq("staff_code", staff_code).execute()
                if existing.data:
                    supabase.table("staff_master").update(data).eq("staff_code", staff_code).execute()
                    updated += 1
                else:
                    supabase.table("staff_master").insert(data).execute()
                    inserted += 1

            except Exception as e:
                errors += 1
                print(f"行のインポートエラー：{row} / {e}")

        # マスタにない職員をis_active=Falseに更新（退職者）
        deactivated = 0
        if new_staff_codes:
            all_staff = supabase.table("staff_master").select("staff_code").eq("is_active", True).execute()
            deactivated = 0
            for s in all_staff.data:
                if s["staff_code"] not in new_staff_codes:
                    supabase.table("staff_master").update({
                        "is_active": False,
                        "updated_at": date.today().isoformat()
                    }).eq("staff_code", s["staff_code"]).execute()
                    deactivated += 1

        send_email(sender_email, "【かりや生協】職員マスタ更新 完了",
            f"""職員マスタの更新が完了しました。

　新規登録：{inserted}名
　更新：{updated}名
　退職・無効化：{deactivated}名
　エラー：{errors}件

更新日：{date.today().strftime('%Y年%m月%d日')}

かりや生協 AIスタッフ""")

        print(f"職員マスタインポート完了：新規{inserted}名、更新{updated}名、無効化{deactivated}名")

    except Exception as e:
        send_email(sender_email, "【かりや生協】職員マスタ更新 - エラー",
            f"インポート中にエラーが発生しました。\n\nエラー内容：{e}\n\nかりや生協 AIスタッフ")
        print(f"職員マスタインポートエラー：{e}")


def handle_work_email_import(sender_email, csv_content):
    """組合メールアドレスをCSVから一括登録する"""
    try:
        reader = csv.reader(io.StringIO(csv_content))
        rows = list(reader)

        if not rows:
            send_email(sender_email, "【かりや生協】組合メール一括登録 - エラー",
                "CSVファイルが空です。確認してください。\n\nかりや生協 AIスタッフ")
            return

        # ヘッダー判定
        first_row = rows[0]
        start_index = 1 if any(h in first_row for h in ["職員コード", "社員CD", "社員No", "メール"]) else 0

        inserted = 0
        updated = 0
        skipped = 0

        for row in rows[start_index:]:
            if len(row) < 2:
                continue
            try:
                staff_code_raw = str(row[0]).strip()
                work_email = str(row[1]).strip()

                if not staff_code_raw.isdigit() or not work_email or "@" not in work_email:
                    skipped += 1
                    continue

                staff_code = int(staff_code_raw)

                existing = supabase.table("email_registry").select("staff_code").eq("staff_code", staff_code).execute()
                if existing.data:
                    supabase.table("email_registry").update({
                        "work_email": work_email,
                        "updated_at": date.today().isoformat()
                    }).eq("staff_code", staff_code).execute()
                    updated += 1
                else:
                    supabase.table("email_registry").insert({
                        "staff_code": staff_code,
                        "work_email": work_email
                    }).execute()
                    inserted += 1

            except Exception as e:
                skipped += 1
                print(f"組合メール登録エラー：{e}")

        send_email(sender_email, "【かりや生協】組合メール一括登録 完了",
            f"""組合メールアドレスの一括登録が完了しました。

　新規登録：{inserted}名
　更新：{updated}名
　スキップ：{skipped}件

更新日：{date.today().strftime('%Y年%m月%d日')}

かりや生協 AIスタッフ""")

        print(f"組合メール一括登録完了：新規{inserted}名、更新{updated}名")

    except Exception as e:
        send_email(sender_email, "【かりや生協】組合メール一括登録 - エラー",
            f"インポート中にエラーが発生しました。\n\nエラー内容：{e}\n\nかりや生協 AIスタッフ")
        print(f"組合メール一括登録エラー：{e}")


def handle_license_import(sender_email, csv_content):
    """免許証データCSVをインポートする"""
    try:
        reader = csv.reader(io.StringIO(csv_content))
        rows = list(reader)

        if not rows:
            send_email(sender_email, "【かりや生協】免許証データ更新 - エラー",
                "CSVファイルが空です。確認してください。\n\nかりや生協 AIスタッフ")
            return

        print(f"免許証CSVの行数：{len(rows)}")

        # 1行目がヘッダーか判定
        first_row = rows[0]
        start_index = 1 if any(h in first_row for h in ["社員No", "社員CD", "職員コード", "有効期限"]) else 0

        inserted = 0
        updated = 0
        skipped = 0
        errors = 0

        for row in rows[start_index:]:
            if len(row) < 2:
                continue
            try:
                # 社員No（列1）と有効期限（列5）を取得
                # フォーマット：所属名, 社員No, 氏名, 氏名カナ, 役職名, 有効期限
                staff_code_raw = str(row[1]).strip()
                if not staff_code_raw or not staff_code_raw.isdigit():
                    skipped += 1
                    continue

                staff_code = int(staff_code_raw)

                # 有効期限を取得（列5、なければ列1の次）
                expiry_raw = row[5].strip() if len(row) > 5 else ""
                if not expiry_raw:
                    skipped += 1
                    continue

                # 日付を解析（各種フォーマット対応）
                expiry_date = None
                for fmt in ["%Y/%m/%d", "%Y-%m-%d", "%Y年%m月%d日", "%Y/%m/%d %H:%M:%S"]:
                    try:
                        from datetime import datetime as dt
                        expiry_date = dt.strptime(expiry_raw.split(" ")[0], fmt).date()
                        break
                    except Exception:
                        continue

                if not expiry_date:
                    skipped += 1
                    continue

                data = {
                    "staff_code": staff_code,
                    "license_expiry_date": expiry_date.isoformat(),
                    "updated_at": date.today().isoformat()
                }

                existing = supabase.table("license_management").select("staff_code").eq("staff_code", staff_code).execute()
                if existing.data:
                    supabase.table("license_management").update(data).eq("staff_code", staff_code).execute()
                    updated += 1
                else:
                    supabase.table("license_management").insert(data).execute()
                    inserted += 1

            except Exception as e:
                errors += 1
                print(f"免許証データ行エラー：{e}")

        send_email(sender_email, "【かりや生協】免許証データ更新 完了",
            f"""免許証データの更新が完了しました。

　新規登録：{inserted}名
　更新：{updated}名
　スキップ：{skipped}件（空欄・ヘッダー等）
　エラー：{errors}件

更新日：{date.today().strftime('%Y年%m月%d日')}

かりや生協 AIスタッフ""")

        print(f"免許証インポート完了：新規{inserted}名、更新{updated}名")

    except Exception as e:
        send_email(sender_email, "【かりや生協】免許証データ更新 - エラー",
            f"インポート中にエラーが発生しました。\n\nエラー内容：{e}\n\nかりや生協 AIスタッフ")
        print(f"免許証インポートエラー：{e}")


def handle_email_registration(sender_email, sender_name, body):
    """メールアドレス登録処理"""
    # 職員コードを本文から抽出
    match = re.search(r'職員コード[：:\s]*(\d+)', body)
    if not match:
        match = re.search(r'社員CD[：:\s]*(\d+)', body)
    if not match:
        match = re.search(r'(\d{4,6})', body)

    if not match:
        send_email(sender_email, "【かりや生協】メールアドレス登録 - 職員コードが見つかりません",
            f"""{sender_name} さん

メールありがとうございます。
本文に職員コードが見つかりませんでした。

以下の形式で再度お送りください：

　職員コード：12345

かりや生協 AIスタッフ""")
        return

    staff_code = int(match.group(1))

    # 職員マスタに存在するか確認
    master = supabase.table("staff_master").select("name").eq("staff_code", staff_code).execute()
    if not master.data:
        send_email(sender_email, "【かりや生協】メールアドレス登録 - 職員コードが見つかりません",
            f"""{sender_name} さん

職員コード「{staff_code}」は職員マスタに登録されていません。
職員コードをご確認の上、再度お送りください。

かりや生協 AIスタッフ""")
        return

    staff_name = master.data[0]["name"]

    # 組合メールか個人メールかを判定
    is_work_email = sender_email.endswith(f"@{WORK_EMAIL_DOMAIN}")
    email_field = "work_email" if is_work_email else "personal_email"
    email_type = "組合メールアドレス" if is_work_email else "個人メールアドレス"

    # メールアドレスを登録・更新
    existing = supabase.table("email_registry").select("staff_code").eq("staff_code", staff_code).execute()
    if existing.data:
        supabase.table("email_registry").update({
            email_field: sender_email,
            "updated_at": date.today().isoformat()
        }).eq("staff_code", staff_code).execute()
        action = "更新"
    else:
        supabase.table("email_registry").insert({
            "staff_code": staff_code,
            email_field: sender_email
        }).execute()
        action = "登録"

    send_email(sender_email, "【かりや生協】メールアドレス登録完了",
        f"""{staff_name} さん

{email_type}の{action}が完了しました。

　職員コード：{staff_code}
　氏名：{staff_name}
　{email_type}：{sender_email}

かりや生協 AIスタッフ""")
    print(f"メールアドレス{action}完了：{staff_name}（{staff_code}）")


def handle_license_update(sender_email, body):
    """免許証更新期限の更新処理"""
    # 職員コードをメールアドレスから逆引き
    # 個人メール・組合メール両方から逆引き
    registry = supabase.table("email_registry").select("staff_code").eq("personal_email", sender_email).execute()
    if not registry.data:
        registry = supabase.table("email_registry").select("staff_code").eq("work_email", sender_email).execute()
    if not registry.data:
        return False

    staff_code = registry.data[0]["staff_code"]
    master = supabase.table("staff_master").select("name").eq("staff_code", staff_code).execute()
    staff_name = master.data[0]["name"] if master.data else "職員"

    # 日付を本文から抽出（例：2027年3月、2027/03、2027-03など）
    match = re.search(r'(\d{4})[年/\-](\d{1,2})', body)
    if not match:
        return False

    year = int(match.group(1))
    month = int(match.group(2))
    # 月末日を有効期限とする
    import calendar
    last_day = calendar.monthrange(year, month)[1]
    expiry_date = date(year, month, last_day)

    existing = supabase.table("license_management").select("staff_code").eq("staff_code", staff_code).execute()
    if existing.data:
        supabase.table("license_management").update({
            "license_expiry_date": expiry_date.isoformat(),
            "updated_at": date.today().isoformat()
        }).eq("staff_code", staff_code).execute()
    else:
        supabase.table("license_management").insert({
            "staff_code": staff_code,
            "license_expiry_date": expiry_date.isoformat()
        }).execute()

    send_email(sender_email, "【かりや生協】免許証更新期限を記録しました",
        f"""{staff_name} さん

免許証の更新期限を記録しました。

　新しい有効期限：{year}年{month}月{last_day}日

ありがとうございました。

かりや生協 AIスタッフ""")
    print(f"免許更新期限を更新：{staff_name}（{staff_code}）→ {expiry_date}")
    return True


def send_license_reminders():
    """免許更新期限が3ヶ月以内の職員にリマインドを送信"""
    try:
        today = date.today()
        three_months_later = today + timedelta(days=90)

        licenses = supabase.table("license_management").select("*").execute()
        reminded_count = 0

        for lic in licenses.data:
            if not lic.get("license_expiry_date"):
                continue

            expiry_date = date.fromisoformat(lic["license_expiry_date"])
            if not (today <= expiry_date <= three_months_later):
                continue

            # メールアドレスを取得（個人メール優先、なければ組合メール）
            registry = supabase.table("email_registry").select("personal_email, work_email").eq("staff_code", lic["staff_code"]).execute()
            if not registry.data:
                continue
            staff_email = registry.data[0]["personal_email"] or registry.data[0]["work_email"]
            if not staff_email:
                continue

            # 職員名を取得
            master = supabase.table("staff_master").select("name").eq("staff_code", lic["staff_code"]).execute()
            staff_name = master.data[0]["name"] if master.data else "職員"
            days_left = (expiry_date - today).days

            send_email(staff_email, "【かりや生協】運転免許証の更新期限のお知らせ",
                f"""{staff_name} さん

お疲れさまです。かりや生協 AIスタッフです。

運転免許証の更新期限が近づいています。

　更新期限：{expiry_date.strftime('%Y年%m月%d日')}（あと{days_left}日）

更新が完了しましたら、このメールに返信する形で
「免許を更新しました。新しい期限は○年○月です。」
とお送りください。自動的に記録を更新します。

かりや生協 AIスタッフ""")

            supabase.table("license_management").update({
                "last_reminded_at": today.isoformat()
            }).eq("staff_code", lic["staff_code"]).execute()

            reminded_count += 1
            print(f"リマインド送信：{staff_name}（{staff_email}）")

        print(f"免許更新リマインド完了：{reminded_count}件")

    except Exception as e:
        print(f"免許更新チェックエラー：{e}")


def generate_reply(sender_name, subject, body):
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"送信者：{sender_name}\n件名：{subject}\n\n本文：\n{body}"}
        ]
    )
    return response.choices[0].message.content


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

            # 組合メール一括登録
            if "組合メール一括登録" in subject:
                csv_content = get_csv_attachment(msg)
                if csv_content:
                    handle_work_email_import(sender_email, csv_content)
                else:
                    send_email(sender_email, "【かりや生協】組合メール一括登録 - CSVが見つかりません",
                        "CSVファイルが添付されていません。\nCSVファイルを添付して再送してください。\n\nかりや生協 AIスタッフ")
                continue

            # 免許証データ更新
            if "免許証データ更新" in subject:
                csv_content = get_csv_attachment(msg)
                if csv_content:
                    handle_license_import(sender_email, csv_content)
                else:
                    send_email(sender_email, "【かりや生協】免許証データ更新 - CSVが見つかりません",
                        "CSVファイルが添付されていません。\nCSVファイルを添付して再送してください。\n\nかりや生協 AIスタッフ")
                continue

            # 職員マスタ更新
            if "職員マスタ更新" in subject:
                csv_content = get_csv_attachment(msg)
                if csv_content:
                    handle_staff_master_import(sender_email, csv_content)
                else:
                    send_email(sender_email, "【かりや生協】職員マスタ更新 - CSVが見つかりません",
                        "CSVファイルが添付されていません。\nCSVファイルを添付して再送してください。\n\nかりや生協 AIスタッフ")
                continue

            # メールアドレス登録
            if "メールアドレス登録" in subject or "メールアドレス登録" in body:
                handle_email_registration(sender_email, sender_name, body)
                continue

            # 免許証更新報告
            if any(kw in body for kw in ["免許を更新", "免許証を更新", "更新しました", "新しい期限"]):
                if handle_license_update(sender_email, body):
                    continue

            # 通常のメール返信
            reply_body = generate_reply(sender_name, subject, body)
            send_email(sender_email, f"Re: {subject}" if not subject.startswith("Re:") else subject, reply_body)
            print(f"返信送信完了: {sender_email}")

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
