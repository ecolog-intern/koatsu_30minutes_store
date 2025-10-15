# tohoku_scraping.py

# tohoku_scraping.py
import os
import shutil
from playwright.async_api import async_playwright
from zip_thawing import ZipThawing

class TohokuScraping:
    def __init__(self, region, config):
        self.region = region
        self.config = config
        self.s3_client = config.s3_client
        self.S3_BUCKET = config.bucket_name
        self.download_dir = config.download_folder
        self.domain = config.tohoku_domain
        self.url = config.tohoku_url
        self.cert_content = config.cert_content
        self.key_content = config.key_content
        self.yesterday_str = config.yesterday_str
        self.year2digit = config.year2digit
        self.month_day = config.month_day

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

            # --- フレーム切り替え（左側メニュー） ---
            print("左側メニュー（contents）フレームへ切り替え中...")
            contents_frame = page.frame(name="contents")
            if not contents_frame:
                print("⚠️ contentsフレームが見つかりません。")
                return

            # 「日毎30分値」をクリック
            links = await contents_frame.query_selector_all("a")
            found = False
            for link in links:
                text = await link.inner_text()
                if "日毎30分値" in text:
                    await link.click()
                    print("✅ 『日毎30分値』をクリックしました。")
                    found = True
                    break

            if not found:
                print("⚠️ 『日毎30分値』リンクが見つかりませんでした。")
                await browser.close()
                return

            await page.wait_for_timeout(8000)

            # --- メインフレームに切り替え ---
            main_frame = page.frame(name="main")
            if not main_frame:
                print("⚠️ mainフレームが見つかりません。")
                await browser.close()
                return

            # 対象ZIPファイル名の組み立て
            target_filename = f"W4012020{self.year2digit}{self.month_day}00000000.zip"
            print(f"[DEBUG] 探索対象ファイル名: {target_filename}")

            # ZIPリンクを探索
            zip_links = await main_frame.query_selector_all('a[href^="javascript:__doPostBack"]')
            found = False
            for link in zip_links:
                text = await link.inner_text()
                if target_filename in text:
                    href = await link.get_attribute("href")
                    event_target = href.split("'")[1]
                    print(f"ダウンロードイベントをトリガー: {event_target}")
                    await main_frame.evaluate(f"__doPostBack('{event_target}', '')")
                    found = True
                    break

            if not found:
                print(f"[WARNING] 対象ファイルが見つかりません: {target_filename}")
                await browser.close()
                return

            # ダウンロード待機
            print("⏳ ダウンロード待機中...")
            async with page.expect_download() as download_info:
                await page.wait_for_timeout(5000)
            download = await download_info.value
            save_path = os.path.join(self.download_dir, download.suggested_filename)
            await download.save_as(save_path)
            print(f"✅ ダウンロード完了: {save_path}")

            # ZIP解凍 → CSV変換
            zipthawing = ZipThawing(self.download_dir)
            csv_path = zipthawing.zip_thawing()

            # ---- S3アップロード ----
            csv_filename = os.path.basename(csv_path)
            s3_key = f"フォルダ保存用/{self.region}/{self.config.yesterday_month_str}/{self.yesterday_str}/{csv_filename}"
            self.s3_client.upload_file(csv_path, self.S3_BUCKET, s3_key)
            print(f"☁️ S3アップロード完了: s3://{self.S3_BUCKET}/{s3_key}")

            # ダウンロードフォルダを削除
            shutil.rmtree(self.download_dir)
            print("🗑️ ローカルファイル削除済み")

            await browser.close()
