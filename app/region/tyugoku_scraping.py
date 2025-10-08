import os
import boto3
import base64
from playwright.async_api import async_playwright

class TyugokuScraping:
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
        self.domain = os.getenv('CHOGOKU_DOMAIN')
        self.url = os.getenv('CHUGOKU_URL')
        

    async def scraping(self):
        os.makedirs(self.download_dir, exist_ok=True)

        async with async_playwright() as p:
            # --- ブラウザ起動（ヘッドレス） ---
            browser = await p.chromium.launch(headless=True)
            context = await browser.new_context(
                accept_downloads=True,
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
                print(f"行 {i}: {text[:100]}")  # 最初の100文字のみ表示
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

                    # ---- S3アップロード ----
                    s3_key = f"フォルダ保存用/{self.region}/{self.yesterday_month_str}/{self.yesterday_str}/{download.suggested_filename}"
                    self.s3_client.upload_file(save_path, self.S3_BUCKET, s3_key)
                    print(f"☁️ S3アップロード完了: s3://{self.S3_BUCKET}/{s3_key}")

                    # ローカル削除（任意）
                    os.remove(save_path)
                    print("🗑️ ローカルファイル削除済み")
                    break

            if not found:
                print("⚠️ 該当データが見つかりませんでした。")

            await browser.close()




