# kanto_scraping.py
import os
import shutil
from playwright.async_api import async_playwright
from zip_thawing import ZipThawing

class KantoScraping:
    def __init__(self, region, config):
        self.region = region
        self.config = config
        self.s3_client = config.s3_client
        self.S3_BUCKET = config.bucket_name
        self.download_dir = config.download_folder
        self.domain = config.kanto_domain
        self.url = config.kanto_url
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
            print("✅ 東京電力サイト（関東）へアクセス成功")

            # --- 「情報公開」ボタンをクリック ---
            print("『情報公開』ボタンをクリック...")
            elem = await page.wait_for_selector("#johokokai", timeout=20000)
            await elem.click()
            await page.wait_for_timeout(5000)

            # --- 「同時同量公開一覧」をクリック ---
            print("『同時同量公開一覧』をクリック...")
            link = await page.wait_for_selector("a:text('同時同量公開一覧')", timeout=20000)
            await link.click()
            await page.wait_for_timeout(5000)

            # --- セレクトボックス設定 ---
            print("データ種別・電圧種別を選択中...")
            await page.select_option("#DTO-LVK4RS001_INFO_KUBUN_CD", "0120")  #日毎同時同量
            await page.select_option("#DTO-LVK4RS001_VOLT_SHUBT_CD", "0")    # 特高・高圧
            await page.wait_for_timeout(1000)

            # --- 検索ボタンをクリック ---
            print("検索実行中...")
            await page.click("[name='j_idt24']")
            await page.wait_for_timeout(3000)

            # --- 最初の結果を選択 ---
            print("一覧から最初の結果を選択...")
            await page.click("[name='DTO-LVK4RS001G02_SELECT_INDEX_LIST[0]']")
            await page.wait_for_timeout(2000)

            # --- ダウンロードボタンをクリック ---
            print("ZIPファイルのダウンロードを開始します...")
            async with page.expect_download() as download_info:
                await page.click("[name='j_idt61']")
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
