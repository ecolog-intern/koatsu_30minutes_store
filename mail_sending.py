'''
outlookメールを送信する
成功時は成功メールを送信して
エラーが出た時はエラーメールを送信する
'''
import pythoncom
import win32com.client

class MailSending:
    def __init__(self, to, subject, body):
        self.to = to
        self.subject = subject
        self.body = body
        
    def send_mail(self):
        #実際にメールを送信する
        pythoncom.CoInitialize()

        # Outlookアプリケーションを起動し、新規メールを作成
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        # メールの宛先、件名、本文を設定
        mail.To = self.to
        mail.Subject = self.subject
        mail.Body = self.body

        # メール送信
        mail.Send()
        print(f"✅ メールを {self.to} に送信しました。")