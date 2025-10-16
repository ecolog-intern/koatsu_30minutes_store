import os
import subprocess
from bs4 import BeautifulSoup

class KansaiScraping:
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

    def scraping(self):
        """
        関西電力サイトにアクセスしてデータを取得
        """
        print("🔗 関西電力サイトにアクセス中...")
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
            "--tlsv1.0",
            "-c", cookie_path,  # Cookie保存
            "-L",
            self.url,
        ]
        subprocess.run(cmd1, capture_output=True, text=False)

        # --- 2️⃣ NSC提供情報メニュー遷移 ---
        print("STEP2: NSC提供情報メニューへ遷移")
        nsc_url = "https://www4.kepco.co.jp/takusouinfo/H24D/H24DF702J.jsp"
        cmd2 = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.0",
            "-b", cookie_path,
            "-c", cookie_path,
            "-L",
            nsc_url,
        ]
        result2 = subprocess.run(cmd2, capture_output=True, text=False)
        html2 = result2.stdout.decode("cp932", errors="replace")
        print("✅ NSC提供情報メニュー中継ページ取得成功")

        # --- 3️⃣ フォーム自動送信（POST） ---
        print("STEP3: フォーム送信で遷移")
        post_url = "https://www4.kepco.co.jp/takusouinfo/H24DF700A04.do"
        cmd3 = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.0",
            "-b", cookie_path,
            "-c", cookie_path,
            "-L",
            "-X", "POST",
            post_url,
        ]
        result3 = subprocess.run(cmd3, capture_output=True, text=False)
        html3 = result3.stdout.decode("cp932", errors="replace")

        # --- 4️⃣ 二重ログイン警告ページを突破 ---
        print("STEP4: 二重ログイン警告突破")
        force_login_url = "https://www4.kepco.co.jp/takusouinfo/H24DF700A03.do"
        cmd4 = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.0",
            "-b", cookie_path,
            "-c", cookie_path,
            "-L",
            "-X", "POST",
            "-d", "actionRequest=Login",
            "-d", "odbango=H24DF701",
            "-d", "updateFlg=",
            "-d", "userId=ESZ773000006",
            "-d", "password=",
            force_login_url,
        ]
        result4 = subprocess.run(cmd4, capture_output=True, text=False)
        html4 = result4.stdout.decode("cp932", errors="replace")
        print("✅ 強制ログイン完了（メニュー画面遷移）")

        with open(os.path.join(self.download_dir, "nsc_force_login.html"), "w", encoding="cp932") as f:
            f.write(html4)

        # --- 5️⃣ 同時同量支援メニューへ ---
        print("STEP5: 同時同量支援・低圧30分値提供ページ遷移")
        target_url = "https://www4.kepco.co.jp/takusouinfo/H24DF710A02.do"
        cmd5 = [
            "curl",
            "--silent", "--show-error",
            "--insecure",
            "--cert", self.cert_path,
            "--key", self.key_path,
            "--tlsv1.0",
            "-b", cookie_path,
            "-c", cookie_path,
            "-L",
            "-X", "POST",
            "-d", "actionRequest=DojiDoryouShien",
            "-d", "odbango=H24DF710",
            "-d", "updateFlg=",
            target_url,
        ]
        result5 = subprocess.run(cmd5, capture_output=True, text=False)
        html5 = result5.stdout.decode("cp932", errors="replace")

        print("✅ ページ取得成功（同時同量支援メニュー）")
        print("=" * 80)
        print(html5[:1500])
        print("=" * 80)

        out_path = os.path.join(self.download_dir, "dojidoryou_page.html")
        with open(out_path, "w", encoding="cp932") as f:
            f.write(html5)
        print(f"📄 HTMLを保存しました: {out_path}")
        print(f"📊 HTML サイズ: {len(html5)} 文字")

        # BeautifulSoupでHTMLを解析
        soup5 = BeautifulSoup(html5, "html.parser")

        # すべてのZIPファイルリンクを抽出
        links = soup5.find_all("a", href="#")
        target_files = []

        import re
        for link in links:
            onclick = link.get("onclick", "")
            if "xmlDL" in onclick:
                filename = link.get_text(strip=True)
                target_date = f"W40120{self.yesterday_str}"
                if filename.endswith(".zip") and target_date in filename:
                    target_files.append((filename, onclick))

        if len(target_files) == 0:
            print("⚠️ 昨日の日付を含むファイルが見つかりませんでした")
            return

        print(f"📊 {len(target_files)} 個のファイルが見つかりました")

        # --- 6️⃣ 特定のファイルをダウンロード ---
        print("\n📥 ダウンロード開始:")
        for fname, onclick in target_files:
            print(f"  📦 {fname}")

            # onclickから各パラメータを抽出
            # 例: xmlDL('0','1','2')
            match = re.search(r"xmlDL\('(\d+)','(\d+)','(\d+)'\)", onclick)
            if match:
                selected_index = match.group(1)
                selected_list = match.group(2)
                koatsu_teatsu_kbn = match.group(3)

                # odbangoの値を取得（H24DF720固定）
                odbango = "H24DF720"

                # ダウンロードURL構築
                download_url = (
                    f"https://www4.kepco.co.jp/takusouinfo/H24DF720A02.do"
                    f"?selectedIndex={selected_index}"
                    f"&selectedList={selected_list}"
                    f"&koatsuTeatsuKbn={koatsu_teatsu_kbn}"
                    f"&actionRequest=Download"
                    f"&odbango={odbango}"
                )

                # curlでダウンロード
                download_path = os.path.join(self.download_dir, fname)
                cmd_download = [
                    "curl",
                    "--silent", "--show-error",
                    "--insecure",
                    "--cert", self.cert_path,
                    "--key", self.key_path,
                    "--tlsv1.0",
                    "-b", cookie_path,
                    "-c", cookie_path,
                    "-L",
                    "-o", download_path,
                    download_url,
                ]

                result_dl = subprocess.run(cmd_download, capture_output=True, text=False)

                if result_dl.returncode == 0:
                    # ファイルサイズを確認
                    if os.path.exists(download_path):
                        file_size = os.path.getsize(download_path)
                        print(f"    ✅ ダウンロード成功: {file_size} bytes")
                    else:
                        print(f"    ❌ ファイルが作成されませんでした")
                else:
                    print(f"    ❌ ダウンロード失敗: {result_dl.stderr.decode('utf-8', errors='replace')}")

        print(f"\n✅ ダウンロード完了: {len(target_files)} ファイル")

        # --- ZIP解凍 → CSV変換 → S3アップロード ---
        if len(target_files) > 0:
            print("\n📦 ZIP解凍とCSV変換を開始...")
            from zip_thawing import ZipThawing
            import shutil

            zipthawing = ZipThawing(self.download_dir)
            csv_path = zipthawing.zip_thawing()
            print(f"✅ CSV変換完了: {csv_path}")

            # --- S3アップロード ---
            csv_filename = os.path.basename(csv_path)
            s3_key = f"フォルダ保存用/{self.region}/{self.yesterday_month_str}/{self.yesterday_str}/{csv_filename}"

            try:
                self.s3_client.upload_file(csv_path, self.S3_BUCKET, s3_key)
                print(f"☁️ S3アップロード完了: s3://{self.S3_BUCKET}/{s3_key}")
            except Exception as e:
                print(f"❌ S3アップロード失敗: {e}")

            # --- ダウンロードフォルダ内のファイルを削除 ---
            for filename in os.listdir(self.download_dir):
                file_path = os.path.join(self.download_dir, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    print(f"⚠️ {file_path} の削除に失敗: {e}")
            print("🗑️ ローカルファイル削除済み")