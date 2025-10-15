# shikoku_scraping.py
import os
import shutil
from datetime import datetime, timedelta
from playwright.async_api import async_playwright
from zip_thawing import ZipThawing

class ShikokuScraping:
    def __init__(self, region, config):
        self.region = region
        self.config = config
        self.s3_client = config.s3_client
        self.S3_BUCKET = config.bucket_name
        self.download_dir = config.download_folder
        self.domain = config.shikoku_domain
        self.url = config.shikoku_url
        self.cert_content = config.cert_content
        self.key_content = config.key_content
        self.yesterday = (datetime.today() - timedelta(days=1))
        self.yesterday_str = self.yesterday.strftime("%Y%m%d")
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
            print("四国電力サイトにアクセス中...")
            await page.goto(self.url, wait_until="networkidle")
            print("✅ アクセス成功:", self.url)

            # --- 「３０分電力量提供システム」をクリック ---
            print("『３０分電力量提供システム』をクリック...")
            await page.click("text=３０分電力量提供システム")
            await page.wait_for_timeout(3000)

            # --- 「30分電力量および日毎30分電力量」をクリック ---
            print("『30分電力量および日毎30分電力量』をクリック...")
            await page.click("text=30分電力量および日毎30分電力量")
            await page.wait_for_timeout(3000)

            # --- 昨日の日付に基づくファイル名を生成 ---
            filename = f"W40120{self.yesterday_str}00000000.zip"
            print(f"探索対象ファイル名: {filename}")

            # --- ファイルリンクを検索してクリック ---
            link = await page.query_selector(f"a:text('{filename}')")
            if not link:
                print("⚠️ ダウンロードリンクが見つかりません。")
                await browser.close()
                return

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

            # --- ローカルフォルダ削除 ---
            shutil.rmtree(self.download_dir)
            print("🗑️ ローカルファイル削除済み")

            await browser.close()
