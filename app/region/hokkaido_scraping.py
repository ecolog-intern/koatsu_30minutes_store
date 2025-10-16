# hokkaido_scraping.py
import os
import subprocess
from datetime import datetime, timedelta

class HokkaidoScraping:
    def __init__(self, region, config):
        self.region = region
        self.config = config
        self.s3_client = config.s3_client
        self.S3_BUCKET = config.bucket_name
        self.download_dir = config.download_folder
        self.domain = config.hokkaido_domain
        self.url = config.hokkaido_url
        self.cert_content = config.cert_content
        self.key_content = config.key_content
        self.yesterday = (datetime.today() - timedelta(days=1)).strftime("%Y%m%d")
        self.yesterday_str = config.yesterday_str
        self.yesterday_month_str = config.yesterday_month_str

        # --- 一時証明書ファイル ---
        self.cert_path = "/tmp/hokkaido_client.crt"
        self.key_path = "/tmp/hokkaido_client.key"
        os.makedirs("/tmp", exist_ok=True)
        with open(self.cert_path, "wb") as f:
            f.write(self.cert_content)
        with open(self.key_path, "wb") as f:
            f.write(self.key_content)

    def scraping(self):
        """
        北海道電力サイトにアクセスしてHTMLレスポンスを表示
        """
        print("🔗 北海道電力サイトにアクセス中...")
        print("= DEBUG CONFIG INFO =")
        print("domain:", self.domain)
        print("url:", self.url)

        os.makedirs(self.download_dir, exist_ok=True)

        # --- curlでアクセス ---
        print("STEP1: 北海道電力URLへアクセス")
        cmd = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.0",
            "-L",
            self.url,
        ]

        result = subprocess.run(cmd, capture_output=True, text=False)

        if result.returncode != 0:
            print(f"❌ curlエラー: {result.stderr.decode('utf-8', errors='replace')}")
            return

        # HTMLを取得（エンコーディングを試行）
        html_content = None
        for encoding in ['utf-8', 'shift_jis', 'euc-jp', 'iso-2022-jp']:
            try:
                html_content = result.stdout.decode(encoding)
                print(f"✅ エンコーディング成功: {encoding}")
                break
            except UnicodeDecodeError:
                continue

        if html_content is None:
            html_content = result.stdout.decode('utf-8', errors='replace')
            print("⚠️ エンコーディング不明。UTF-8でエラーを無視して表示")

        # HTMLを表示
        print("=" * 80)
        print("最初の画面のHTML:")
        print("=" * 80)
        print(html_content)
        print("=" * 80)
        print(f"📊 HTML サイズ: {len(html_content)} 文字")

        # HTMLファイルを保存
        out_path = os.path.join(self.download_dir, "hokkaido_page.html")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        print(f"📄 HTMLを保存しました: {out_path}")
