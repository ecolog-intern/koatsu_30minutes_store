import os
import asyncio
import shutil
from bs4 import BeautifulSoup
from zip_thawing import ZipThawing

class ShikokuScraping:
    def __init__(self, region, config):
        """
        Configクラスの設定を受け取って利用する構成
        """
        self.region = region
        self.config = config
        self.s3_client = config.s3_client
        self.S3_BUCKET = config.bucket_name
        self.download_dir = config.download_folder
        self.domain = config.kansai_domain
        self.url = config.kansai_url
        self.cert_content = config.cert_content
        self.key_content = config.key_content
        self.yesterday_str = config.yesterday_str
        self.yesterday_month_str = config.yesterday_month_str

        # --- 一時証明書ファイル ---
        self.cert_path = "/tmp/client.crt"
        self.key_path = "/tmp/client.key"
        os.makedirs("/tmp", exist_ok=True)
        with open(self.cert_path, "wb") as f:
            f.write(self.cert_content)
        with open(self.key_path, "wb") as f:
            f.write(self.key_content)

    async def _run_curl(self, cmd):
        """非同期でcurlを実行"""
        process = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        stdout, stderr = await process.communicate()
        if process.returncode != 0:
            print(f"⚠️ curl実行エラー: {stderr.decode(errors='ignore')}")
        return stdout.decode("cp932", errors="replace")

    async def scraping(self):
        """
        四国電力サイトにアクセスしてデータを取得
        """
        print("🔗 四国電力サイトにアクセス中...")
        os.makedirs(self.download_dir, exist_ok=True)
        cookie_path = "/tmp/kepco_cookie.txt"

        # --- 1️⃣ トップページ ---
        print("STEP1: トップページへアクセス")
        cmd1 = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.2",
            "-c", cookie_path,
            "-L",
            self.url,
        ]
        await self._run_curl(cmd1)

        # --- 2️⃣ 託送関連業務メニュー ---
        print("STEP2: 託送関連業務メニュー")
        nsc_url = "http://www.w3.org/1999/xhtml"
        cmd2 = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.2",
            "-b", cookie_path,
            "-c", cookie_path,
            "-L",
            nsc_url,
        ]
        html2 = await self._run_curl(cmd2)
        print("✅ 託送関連業務メニューページ取得成功")

        print("STEP3: サブメニューへ遷移")
        post_url = "https://wsc3.yonden.co.jp/PPS/Menu1.action"
        cmd3 = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.2",
            "-b", cookie_path,
            "-c", cookie_path,
            "-L",
            "-X", "POST",
            post_url,
        ]
        html3 = await self._run_curl(cmd3)
        print("✅ サブメニューページ取得成功")

        # --- 4️⃣ 受信可能ファイル一覧 ---
        print("STEP4: 受信可能ファイル一覧ページへ")
        post_url = "https://wsc3.yonden.co.jp/PPS/30min/FileListReceiver.action"
        cmd4 = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.2",
            "-b", cookie_path,
            "-c", cookie_path,
            "-L",
            "-X", "POST",
            post_url,
        ]
        html4 = await self._run_curl(cmd4)
        print("✅ 受信可能ファイル一覧ページ取得成功")

        # --- 5️⃣ ZIP存在確認 ---
        filename = f"W40120{self.yesterday_str}00000000.zip"
        print(f"探索対象ファイル名: {filename}")

        soup = BeautifulSoup(html4, "html.parser")
        link_found = any(filename in (a.get("href") or "") for a in soup.find_all("a"))
        if not link_found:
            print(f"⚠️ 対象ファイル（{filename}）は存在しません。処理をスキップします。")
            return
        print("✅ 対象ZIPの存在を確認しました。")

        # --- 6️⃣ データ取得 ---
        print("STEP6: データ取得")
        zip_url = f"https://wsc3.yonden.co.jp/PPS/30min/FileReceiver.action?file={filename}"
        zip_path = os.path.join(self.download_dir, filename)

        print(f"💡 ダウンロードURL: {zip_url}")
        cmd5 = [
            "curl",
            "--silent",
            "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.2",
            "-b", cookie_path,
            "-L",
            "-o", zip_path,
            zip_url
        ]
        await self._run_curl(cmd5)
        print(f"✅ ZIPダウンロード完了: {zip_path}")

        if not os.path.exists(zip_path) or os.path.getsize(zip_path) == 0:
            print("❌ ZIPファイルのダウンロードに失敗しました。")
            return

        # --- 7️⃣ ZIP解凍 → CSV ---
        zipthawing = ZipThawing(self.download_dir)
        csv_path = zipthawing.zip_thawing()
        print("✅ ZIP解凍完了")

        # --- 8️⃣ S3アップロード ---
        csv_filename = os.path.basename(csv_path)
        s3_key = f"フォルダ保存用/{self.region}/{self.yesterday_month_str}/{self.yesterday_str}/{csv_filename}"
        self.s3_client.upload_file(csv_path, self.S3_BUCKET, s3_key)
        print(f"☁️ S3アップロード完了: s3://{self.S3_BUCKET}/{s3_key}")

        # --- 9️⃣ ローカル削除 ---
        shutil.rmtree(self.download_dir)
        print("🗑️ ローカルファイル削除済み")
