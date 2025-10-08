import os
import boto3
import base64
from playwright.async_api import async_playwright

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
        self.domain = os.getenv('KANSAI_DOMAIN') #ここ変える
        self.url = os.getenv('KANSAI_URL') # ここ変える

    async def scraping(self):
        os.makedirs(self.download_dir, exist_ok=True)

        async with async_playwright() as p:
            # --- ブラウザ起動（ヘッドレス） ---
            # Firefoxを使用(古いSSL/TLSサーバーとの互換性のため)
            browser = await p.chromium.launch(
                headless=True,
                args=[
                    '--ignore-certificate-errors',
                    '--disable-web-security',
                    '--allow-insecure-localhost',
                    '--enable-features=LegacyRenegotiation'  # これが重要
                ]
            )
            context = await browser.new_context(
                accept_downloads=True,
                ignore_https_errors=True,  # ← これが重要
                client_certificates=[
                    {
                        "origin": self.domain,
                        "cert": self.cert_content,
                        "key": self.key_content
                    }
                ]
            )
            page = await context.new_page()

            print("URLにアクセス中...")
            await page.goto(self.url, wait_until="networkidle")

            # 最初の画面のHTMLを表示
            html_content = await page.content()
            print("=" * 80)
            print("最初の画面のHTML:")
            print("=" * 80)
            print(html_content)
            print("=" * 80)

            await browser.close()