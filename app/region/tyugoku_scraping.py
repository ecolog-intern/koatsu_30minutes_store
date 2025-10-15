# tyugoku_scraping.py
import os
import shutil
from playwright.async_api import async_playwright
from zip_thawing import ZipThawing

class TyugokuScraping:
    def __init__(self, region, config):
        self.region = region
        self.config = config
        self.s3_client = config.s3_client
        self.S3_BUCKET = config.bucket_name
        self.download_dir = config.download_folder
        self.domain = config.tyugoku_domain
        self.url = config.tyugoku_url
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

            print("URLにアクセス中...")
            await page.goto(self.url, wait_until="networkidle")

            # --- メニュークリック ---
            print("メニューを開いています...")
            pull_menu = await page.wait_for_selector("#pullMenuDiv")
            menu_links = await pull_menu.query_selector_all("a")

            if len(menu_links) > 1:
                await menu_links[1].click()
                await page.wait_for_timeout(3000)
            else:
                print("メニューリンクが見つかりません。")

            print(f"対象日付: {self.yesterday_str}")

            # --- テーブル探索 ---
            print("テーブルを探索中...")
            content_table = await page.wait_for_selector("#contentBody1DetailsElement")
            rows = await content_table.query_selector_all("tr")

            print(f"テーブル行数: {len(rows)}")
            found = False
            for i, row in enumerate(rows):
                text = await row.inner_text()
                if "特高・高圧日毎３０分電力量" in text and self.yesterday_str in text:
                    print("対象行を発見、ファイルをクリックします。")
                    file_link = await row.query_selector("a")
                    await file_link.scroll_into_view_if_needed()
                    async with page.expect_download() as download_info:
                        await file_link.click()
                    download = await download_info.value
                    save_path = os.path.join(self.download_dir, download.suggested_filename)
                    await download.save_as(save_path)
                    print(f"✅ ダウンロード完了: {save_path}")
                    found = True

                    # zip解凍 → csv
                    zipthwing = ZipThawing(self.download_dir)
                    csv_path = zipthwing.zip_thawing()

                    # ---- S3アップロード ----
                    csv_filename = os.path.basename(csv_path)
                    s3_key = f"フォルダ保存用/{self.region}/{self.yesterday_month_str}/{self.yesterday_str}/{csv_filename}"
                    self.s3_client.upload_file(csv_path, self.S3_BUCKET, s3_key)
                    print(f"☁️ S3アップロード完了: s3://{self.S3_BUCKET}/{s3_key}")

                    shutil.rmtree(self.download_dir)
                    print("🗑️ ローカルファイル削除済み")
                    break

            if not found:
                print("⚠️ 該当データが見つかりませんでした。")

            await browser.close()
