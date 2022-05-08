# SoftBank クリエイティブの本にある サンプルコードです。
import sys
import openpyxl
from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Excelの読み込み
wb_master = openpyxl.load_workbook("Excel.xlsx")
ws_master = wb_master["Sheet1"]

customer_list = []
for row in ws_master.iter_rows(min_row=2):
  if row[0].value is None:
    break
  value_list = []
  for c in row:
    value_list.append(c.value)
  customer_list.append(value_list)

# 請求書PDFのフォルダー
pdf_dir = "請求書PDF_202007"

# メーリングリスト
mailing_list = []

# フォルダーから請求書のPDFファイルを1つずつ取得する
for invoice in Path(pdf_dir).glob("*.pdf"):
  # 「顧客ID」 は、 PDFファイルの拡張子を除いた部分
  customer_id = invoice.stem
  # 該当する顧客データを「Excel」から検索
  for customer in customer_list:
    if customer_id == customer[0]:
      # メーリングリストに「顧客データ」と
      # 「PDFファイル」のパスを追加
      mailing_list.append([customer, invoice])
      # 顧客ID、顧客名、メールアドレス、
      # PDFファイルのパスを表示
      print(customer[0], customer[1], customer[4], invoice)

# モード選択
print()  # 1行空ける
mode = input("モード選択（テスト＝test、本番=real）：")
# 本番以外はテスト
if mode != "real":
  test_mode = True
else:
  test_mode = False

# 送信確認
if test_mode:
  result = input("テストモードで自分宛てに送信します（続行＝yes、中止＝no）：")
else:
  result = input("本番モードで送信します（続行＝yes、中止＝no）：")

# 続行以外は中止
if result != "yes":
  print("プログラムを中止します")
  sys.exit()

# メール本文をファイルから読み込んでおく
text = open("mail_body.txt", encoding="utf-8")
body_temp = text.read()
text.close()

my_address = "自分のメールアドレス"

# SMTPサーバー（今回はGmailで送信）
smtp_server = "smtp.gmail.com"
port_number = 587

# ログイン情報（今回はGmailのアカウントを入力する）
account = "自分のメールのアカウント"
password = "自分のメールのパスワード"

# SMTPサーバーに接続
server = smtplib.SMTP(smtp_server, port_number)
server.starttls()
server.login(account, password)

# メーリングリストの顧客に1つずつメール送信
for data in mailing_list:
  customer = data[0]
  pdf_file = data[1]

  # メッセージの準備
  msg = MIMEMultipart()
  # 件名、メールアドレスの設定
  msg["Subject"] = "Excel送付テスト"
  msg["From"] = my_address
  if test_mode:
      msg["To"] = my_address
  else:
      msg["To"] = customer[4]

  # メール本文の追加
  body_text = body_temp.format(
    company=customer[1],
    department=customer[2],
    person=customer[3]
  )
  body = MIMEText(body_text)
  msg.attach(body)
  # 添付ファイルの追加
  pdf = open(pdf_file, mode="rb")
  pdf_data = pdf.read()
  pdf.close()
  attach_file = MIMEApplication(pdf_data)
  attach_file.add_header("Content-Disposition",
                          "attachment", filename=pdf_file.name)
  msg.attach(attach_file)

  # メール送信
  print("メール送信：", customer[0], customer[1])
  server.send_message(msg)

# SMTPサーバーとの接続を閉じる
server.quit()
