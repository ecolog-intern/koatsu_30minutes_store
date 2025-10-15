# tyubu_scraping.py
import os
import shutil
from playwright.async_api import async_playwright
from zip_thawing import ZipThawing

class TyubuScraping:
    def __init__(self, region, config):
        self.region = region
        self.config = config
        self.s3_client = config.s3_client
        self.S3_BUCKET = config.bucket_name
        self.download_dir = config.download_folder
        self.domain = config.tyubu_domain
        self.url = config.tyubu_url
        self.cert_content = config.cert_content
        self.key_content = config.key_content
        self.yesterday_str = config.yesterday_str
        self.yesterday_month_str = config.yesterday_month_str

    async def scraping(self):
        os.makedirs(self.download_dir, exist_ok=True)
        print("= DEBUG CONFIG INFO =")
        print("domain:", self.domain)
        print("url:", self.url)

        async with async_playwright() as p:
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

            print("URLにアクセス中...")
            await page.goto(self.url, wait_until="networkidle")
            print("✅ 中部電力サイトへアクセス成功")

            # --- 「menu」フレームに切り替え ---
            print("メニューフレームに切り替え中...")
            menu_frame = page.frame(name="menu")
            if not menu_frame:
                print("⚠️ menuフレームが見つかりません。")
                await browser.close()
                return

            # 「日毎３０分電力量」をクリック
            links = await menu_frame.query_selector_all("a")
            found = False
            for link in links:
                text = await link.inner_text()
                if "日毎３０分電力量" in text:
                    await link.click()
                    print("✅ 『日毎３０分電力量』をクリックしました。")
                    found = True
                    break

            if not found:
                print("⚠️ 『日毎３０分電力量』リンクが見つかりません。")
                await browser.close()
                return

            await page.wait_for_timeout(4000)

            # --- 「main」フレームに切り替え ---
            print("メインフレームに切り替え中...")
            main_frame = page.frame(name="main")
            if not main_frame:
                print("⚠️ mainフレームが見つかりません。")
                await browser.close()
                return

            # --- ダウンロードリンク探索 ---
            print("ダウンロードリンクを探索中...")
            links = await main_frame.query_selector_all("a")

            if not links:
                print("⚠️ ダウンロードリンクが見つかりません。")
                await browser.close()
                return

            # 1つ目のリンクをクリックしてダウンロード
            async with page.expect_download() as download_info:
                await links[0].click()
                print("ダウンロード開始を待機中...")

            download = await download_info.value
            save_path = os.path.join(self.download_dir, download.suggested_filename)
            await download.save_as(save_path)
            print(f"✅ ダウンロード完了: {save_path}")

            # --- ZIP解凍 → CSV変換 ---
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
