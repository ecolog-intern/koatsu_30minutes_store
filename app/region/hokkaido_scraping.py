# hokkaido_scraping.py
import os
import shutil
from datetime import datetime, timedelta
from playwright.async_api import async_playwright
from zip_thawing import ZipThawing

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

    async def scraping(self):
        os.makedirs(self.download_dir, exist_ok=True)
        print("= DEBUG CONFIG INFO =")
        print("domain:", self.domain)
        print("url:", self.url)

        async with async_playwright() as p:
            # --- ブラウザ起動 ---
            browser = await p.chromium.launch(headless=True)
            context = await browser.new_context(
                accept_downloads=True,
                client_certificates=[
                    {
                        "origin": self.domain,
                        "cert": self.cert_content,
                        "key": self.key_content,
                    }
                ],
            )
            page = await context.new_page()

            # --- サイトアクセス ---
            print("北海道電力サイトにアクセス中...")
            await page.goto(self.url, wait_until="networkidle")
            print("✅ アクセス成功:", self.url)

            current_url = page.url
            print(f"🌐 現在のURL: {current_url}")

            print("同時同量公開のクリック")
            await page.evaluate("""
            const el = [...document.querySelectorAll('a')].find(a => a.textContent.includes('同時同量公開'));
            if (el) el.click();
            """)
            print("✅ JSレベルで『同時同量公開』をクリックしました。")

            current_url = page.url
            print(f"🌐 現在のURL: {current_url}")

            # --- 昨日の日付を入力 ---
            print(f"日付入力: {self.yesterday}")
            await page.fill("#targetDateFrom", self.yesterday)
            await page.fill("#targetDateTo", self.yesterday)

            # --- 表示ボタンをクリック ---
            print("表示ボタンをクリック...")
            await page.click("#displayBtn")
            await page.wait_for_timeout(5000)

            # --- ZIPリンク探索 ---
            print("ダウンロードリンクを探索中...")
            link = await page.query_selector("//td[@class='title']/a")
            if not link:
                print("⚠️ ダウンロードリンクが見つかりません。")
                await browser.close()
                return

            # --- ZIPファイルをダウンロード ---
            async with page.expect_download() as download_info:
                await link.click()
                print("ダウンロード開始を待機中...")
            download = await download_info.value

            save_path = os.path.join(self.download_dir, download.suggested_filename)
            await download.save_as(save_path)
            print(f"✅ ダウンロード完了: {save_path}")

            # --- ZIP解凍 → CSV変換 ---
            print("ZIPを展開し、CSVに変換中...")
            zipthawing = ZipThawing(self.download_dir)
            csv_path = zipthawing.zip_thawing()

            # --- S3アップロード ---
            csv_filename = os.path.basename(csv_path)
            s3_key = f"フォルダ保存用/{self.region}/{self.yesterday_month_str}/{self.yesterday_str}/{csv_filename}"
            self.s3_client.upload_file(csv_path, self.S3_BUCKET, s3_key)
            print(f"☁️ S3アップロード完了: s3://{self.S3_BUCKET}/{s3_key}")

            # --- ダウンロードフォルダ削除 ---
            shutil.rmtree(self.download_dir)
            print("🗑️ ローカルファイル削除済み")

            await browser.close()
