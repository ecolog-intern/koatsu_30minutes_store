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

        # curlコマンドを構築（古いTLS対応のためOpenSSL利用）
        cmd = [
            "curl",
            "--silent", "--show-error",
            "--insecure",              # 証明書検証をスキップ（TLSエラー回避）
            "--cert", self.cert_path,  # クライアント証明書
            "--key", self.key_path,    # クライアント秘密鍵
            "--tlsv1.0",
            "-L",                      # ← 追加: リダイレクトを自動追跡
            self.url
        ]

        # 実行
        result = subprocess.run(cmd, capture_output=True, text=False)

        if result.returncode != 0:
            print("❌ 通信エラー:", result.stderr.decode("utf-8", errors="ignore"))
            return

        # 明示的にWindows-31J(cp932)でデコード
        html_bytes = result.stdout
        html_content = html_bytes.decode("cp932", errors="replace")

        print("✅ HTML取得成功")
        print("=" * 80)
        print("最初の画面のHTML:")
        print("=" * 80)
        print(html_content[:2000])  # 長すぎる場合は冒頭だけ表示
        print("=" * 80)

        # BeautifulSoupで解析
        soup = BeautifulSoup(html_content, "html.parser")
        title = soup.title.string if soup.title else "タイトル不明"
        print(f"🧩 ページタイトル: {title}")

        # HTMLをファイル保存（テキストモードで書く）
        file_path = os.path.join(self.download_dir, "first_page.html")
        with open(file_path, "w", encoding="cp932", errors="replace") as f:
            f.write(html_content)
        print(f"📄 HTMLを {file_path} に保存しました。")
