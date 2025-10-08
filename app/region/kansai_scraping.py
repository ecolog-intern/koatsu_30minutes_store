import os
import boto3
import base64
from playwright.async_api import async_playwright
import subprocess
from bs4 import BeautifulSoup

class KansaiScraping:
    def __init__(self, region, yesterday_month_str, yesterday_str):
        self.region = region
        self.yesterday_month_str = yesterday_month_str
        self.yesterday_str = yesterday_str
        self.S3_BUCKET = os.getenv('bucket_name')
        self.S3_PREFIX = f"tyugoku/{yesterday_month_str}/"
        self.AWS_ACCESS_KEY = os.getenv('AWS_ACCESS_KEY_ID')
        self.AWS_SECRET_KEY = os.getenv('AWS_SECRET_ACCESS_KEY')
        self.AWS_REGION = os.getenv('region')
        client_cert_base64 = os.getenv('CLIENT_CERT_BASE64')
        client_key_base64 = os.getenv('CLIENT_KEY_BASE64')
        self.cert_content = base64.b64decode(client_cert_base64)
        self.key_content = base64.b64decode(client_key_base64)
        self.download_dir = "downloads"
        self.s3_client = boto3.client(
            "s3",
            aws_access_key_id=self.AWS_ACCESS_KEY,
            aws_secret_access_key=self.AWS_SECRET_KEY,
            region_name=self.AWS_REGION,
        )
        self.domain = os.getenv('KANSAI_DOMAIN')
        self.url = os.getenv('KANSAI_URL')

        # 一時証明書ファイル
        self.cert_path = "/tmp/client.crt"
        self.key_path = "/tmp/client.key"
        os.makedirs("/tmp", exist_ok=True)
        with open(self.cert_path, "wb") as f:
            f.write(self.cert_content)
        with open(self.key_path, "wb") as f:
            f.write(self.key_content)

    def scraping(self):
        print("🔗 関西電力サイトにアクセス中...")
        os.makedirs(self.download_dir, exist_ok=True)

        cookie_path = "/tmp/kepco_cookie.txt"

        # --- 1️⃣ トップページ ---
        cmd = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.0",
            "-c", cookie_path,   # ← Cookie保存
            "-L",
            self.url
        ]
        subprocess.run(cmd, capture_output=True, text=False)

        # --- 2️⃣ NSC提供情報メニュー遷移 ---
        nsc_url = "https://www4.kepco.co.jp/takusouinfo/H24D/H24DF702J.jsp"
        cmd2 = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.0",
            "-b", cookie_path,   # ← Cookie再利用
            "-c", cookie_path,   # ← Cookie更新保存
            "-L",
            nsc_url
        ]
        result2 = subprocess.run(cmd2, capture_output=True, text=False)
        html2 = result2.stdout.decode("cp932", errors="replace")

        print("✅ ＮＳＣ提供情報メニュー中継ページ取得成功")

        # --- 3️⃣ フォーム自動送信（POST） ---
        post_url = "https://www4.kepco.co.jp/takusouinfo/H24DF700A04.do"
        print(f"➡ フォーム送信（POST）: {post_url}")

        cmd3 = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.0",
            "-b", cookie_path,
            "-c", cookie_path,
            "-L",
            "-X", "POST",
            post_url
        ]
        result3 = subprocess.run(cmd3, capture_output=True, text=False)
        html3 = result3.stdout.decode("cp932", errors="replace")


        # --- 4️⃣ 二重ログイン警告ページを強制的に突破 ---
        force_login_url = "https://www4.kepco.co.jp/takusouinfo/H24DF700A03.do"
        print(f"⚠️ 二重ログイン警告を検知。強制ログインします: {force_login_url}")

        cmd4 = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.0",
            "-b", "/tmp/kepco_cookie.txt",
            "-c", "/tmp/kepco_cookie.txt",
            "-L",
            "-X", "POST",
            "-d", "actionRequest=Login",
            "-d", "odbango=H24DF701",
            "-d", "updateFlg=",
            "-d", "userId=ESZ773000006",
            "-d", "password=",
            force_login_url
        ]
        result4 = subprocess.run(cmd4, capture_output=True, text=False)
        html4 = result4.stdout.decode("cp932", errors="replace")

        print("✅ 強制ログイン完了（実際のメニュー画面に遷移）")
        print("=" * 80)
        print(html4)
        print("=" * 80)

        # 保存
        with open(os.path.join(self.download_dir, "nsc_force_login.html"), "w", encoding="cp932") as f:
            f.write(html4)


        # --- 5️⃣ 「同時同量支援・低圧30分値提供」ボタンを押す ---
        print("➡ 『同時同量支援・低圧30分値提供』に遷移します...")

        target_url = "https://www4.kepco.co.jp/takusouinfo/H24DF710A02.do"
        cmd5 = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.0",
            "-b", "/tmp/kepco_cookie.txt",
            "-c", "/tmp/kepco_cookie.txt",
            "-L",
            "-X", "POST",
            "-d", "actionRequest=DojiDoryouShien",
            "-d", "odbango=H24DF710",
            "-d", "updateFlg=",
            target_url
        ]

        result5 = subprocess.run(cmd5, capture_output=True, text=False)
        html5 = result5.stdout.decode("cp932", errors="replace")

        print("✅ ページ取得成功（同時同量支援メニュー）")
        print("=" * 80)
        print(html5[:1500])
        print("=" * 80)

        with open(os.path.join(self.download_dir, "dojidoryou_page.html"), "w", encoding="cp932") as f:
            f.write(html5)
        print("📄 HTMLを保存しました: downloads/dojidoryou_page.html")

