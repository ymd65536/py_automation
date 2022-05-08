# SoftBank クリエイティブの本にある サンプルコードです。

import smtplib

# SMTPサーバー（今回はGmailで送信）
smtp_server = "smtp.gmail.com"
port_number = 587

# ログイン情報（今回はGmailのアカウントを入力する）
account = "自分のメールのアカウント"
password = "自分のメールのパスワード"

# 1）SMTPサーバーの指定
server = smtplib.SMTP(smtp_server, port_number)

# SMTPサーバーの応答確認
res_server = server.noop()
print(res_server)

# 2）暗号化通信の開始
res_starttls = server.starttls()
print(res_starttls)

# 3）ログイン
res_login = server.login(account, password)
print(res_login)

# 5）接続を閉じる
server.quit()
