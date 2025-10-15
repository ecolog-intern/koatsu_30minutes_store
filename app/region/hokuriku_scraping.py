# hokuriku_scraping.py

# hokuriku_scraping.py
import os
import shutil
from playwright.async_api import async_playwright
from zip_thawing import ZipThawing

class HokurikuScraping:
    def __init__(self, region, config):
        self.region = region
        self.config = config
        self.s3_client = config.s3_client
        self.S3_BUCKET = config.bucket_name
        self.download_dir = config.download_folder
        self.domain = config.hokuriku_domain
        self.url = config.hokuriku_url
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
                ]
            )
            page = await context.new_page()

            # --- サイトアクセス ---
            print("URLにアクセス中...")
            await page.goto(self.url, wait_until="networkidle")
            print("✅ 北陸電力サイトへアクセス成功")

            # --- 「同時同量支援メッセージ公開」をクリック ---
            print("『同時同量支援メッセージ公開』をクリックします...")
            link = await page.wait_for_selector("text=同時同量支援メッセージ公開", timeout=15000)
            await link.click()
            await page.wait_for_timeout(1500)

            # --- TFBO100_S041ボタンをクリック ---
            print("ボタンをクリック中...")
            button = await page.wait_for_selector("#TFBO100_S041", timeout=10000)
            await button.click()
            await page.wait_for_timeout(3000)

            # --- 「日毎30分同時同量支援メッセージ」セクションを検索 ---
            print("日毎30分同時同量支援メッセージ セクションを探索中...")
            await page.wait_for_selector("text=日毎30分同時同量支援メッセージ", timeout=15000)
            ancestor = await page.query_selector('//div[text()="日毎30分同時同量支援メッセージ"]/ancestor::*[4]')
            tbody = await ancestor.query_selector("tbody")
            first_row = await tbody.query_selector("tr")
            await first_row.click()
            await page.wait_for_timeout(2000)

            # --- ダウンロードボタンをクリック ---
            print("ダウンロードボタンをクリックします...")
            dl_button = await page.query_selector(
                "#shoResultHigoto30minDojiDoryoShienMsgHigoto30minDojiDoryoShienMsgDlBtnJodan"
            )
            async with page.expect_download() as download_info:
                await dl_button.click()
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

            # --- ローカル削除 ---
            shutil.rmtree(self.download_dir)
            print("🗑️ ローカルファイル削除済み")

            await browser.close()
