# SoftBank クリエイティブの本にある サンプルコードです。

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# 自分のメールアドレス
my_address = "自分のメールアドレス"

# SMTPサーバー（今回はGmailで送信）
smtp_server = "smtp.gmail.com"
port_number = 587

# ログイン情報（今回はGmailのアカウントを入力する）
account = "自分のメールのアカウント"
password = "自分のメールのパスワード"

# メッセージの準備
msg = MIMEMultipart()

# 件名、メールアドレスの設定
msg["Subject"] = "送付テスト"
msg["From"] = my_address
msg["To"] = my_address

# メール本文の追加
text = open("mail_body.txt", encoding="utf-8")
body_temp = text.read()
text.close()
body_text = body_temp.format(
    company="株式会社あああ",
    department="Excel部",
    person="Kento.Yamada"
)
body = MIMEText(body_text)
msg.attach(body)

# 添付ファイルの追加
pdf = open("なんでもいいので.pdf", mode="rb")
pdf_data = pdf.read()
pdf.close()
attach_file = MIMEApplication(pdf_data)
attach_file.add_header("Content-Disposition","attachment",filename="なんでもいいので.pdf")
msg.attach(attach_file)

# SMTPサーバーに接続
server = smtplib.SMTP(smtp_server, port_number)
server.starttls()
server.login(account, password)

# メール送信
server.send_message(msg)

# SMTPサーバーとの接続を閉じる
server.quit()
