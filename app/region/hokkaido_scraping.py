# hokkaido_scraping.py
import os
import subprocess
from datetime import datetime, timedelta

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
        北海道電力サイトにアクセスしてデータを取得
        """
        print("🔗 北海道電力サイトにアクセス中...")
        os.makedirs(self.download_dir, exist_ok=True)
        cookie_path = "/tmp/cookie.txt"

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
        result1 = subprocess.run(cmd1, capture_output=True, text=False)

        if result1.returncode != 0:
            print(f"❌ curlエラー: {result1.stderr.decode('utf-8', errors='replace')}")
            return

        # HTMLを取得（エンコーディングを試行）
        html1 = None
        for encoding in ['utf-8', 'shift_jis', 'euc-jp', 'iso-2022-jp']:
            try:
                html1 = result1.stdout.decode(encoding)
                print(f"✅ エンコーディング成功: {encoding}")
                break
            except UnicodeDecodeError:
                continue

        if html1 is None:
            html1 = result1.stdout.decode('utf-8', errors='replace')
            print("⚠️ エンコーディング不明。UTF-8でエラーを無視して表示")

        print("✅ ページ取得成功（トップページ）")
        print("=" * 80)
        print(html1)
        print("=" * 80)

        # HTMLファイルを保存
        out_path = os.path.join(self.download_dir, "hokkaido_page.html")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html1)
        print(f"📄 HTMLを保存しました: {out_path}")
        print(f"📊 HTML サイズ: {len(html1)} 文字")

        # --- 2️⃣ トークンを抽出 ---
        print("\nSTEP2: トークンを抽出")
        import re
        one_time_token_match = re.search(r'id="oneTimeToken" value="([^"]+)"', html1)
        login_token_match = re.search(r'id="loginToken" value="([^"]+)"', html1)

        if one_time_token_match and login_token_match:
            one_time_token = one_time_token_match.group(1)
            login_token = login_token_match.group(1)
            print(f"✅ oneTimeToken: {one_time_token[:20]}...")
            print(f"✅ loginToken: {login_token[:20]}...")
        else:
            print("❌ トークンが見つかりませんでした")
            return

        # --- 3️⃣ 同時同量公開ページへ遷移 ---
        print("\nSTEP3: 同時同量公開ページへ遷移（POSTリクエスト）")
        doujidouryou_url = f"{self.domain}/LNXWPWSS06OH/GPIpG9040"

        # POSTリクエストでトークンを送信
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
            "-X", "POST",
            "-H", "Content-Type: application/x-www-form-urlencoded",
            "-H", f"Referer: {doujidouryou_url}",
            "-H", "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "-H", "Origin: https://nsc.hepco.co.jp",
            "-H", "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36",
            "-d", f"loginToken={login_token}",
            "-d", f"oneTimeToken={one_time_token}",
            doujidouryou_url,
        ]
        result2 = subprocess.run(cmd2, capture_output=True, text=False)

        if result2.returncode != 0:
            print(f"❌ curlエラー: {result2.stderr.decode('utf-8', errors='replace')}")
            return

        # HTMLを取得（エンコーディングを試行）
        html2 = None
        for encoding in ['utf-8', 'shift_jis', 'euc-jp', 'iso-2022-jp']:
            try:
                html2 = result2.stdout.decode(encoding)
                print(f"✅ エンコーディング成功: {encoding}")
                break
            except UnicodeDecodeError:
                continue

        if html2 is None:
            html2 = result2.stdout.decode('utf-8', errors='replace')
            print("⚠️ エンコーディング不明。UTF-8でエラーを無視して表示")

        print("✅ ページ取得成功（同時同量公開）")
        print("=" * 80)
        print(html2[:1500])
        print("=" * 80)

        # HTMLファイルを保存
        out_path2 = os.path.join(self.download_dir, "hokkaido_doujidouryou_page.html")
        with open(out_path2, "w", encoding="utf-8") as f:
            f.write(html2)
        print(f"📄 HTMLを保存しました: {out_path2}")
        print(f"📊 HTML サイズ: {len(html2)} 文字")

        # --- 4️⃣ 日付と電圧区分を設定して表示 ---
        print("\nSTEP4: 日付と電圧区分を設定して表示ボタンを押す")

        # 昨日の日付をYYYY/MM/DD形式に変換
        target_date = self.yesterday_str  # YYYYMMDD形式
        target_date_formatted = f"{target_date[:4]}/{target_date[4:6]}/{target_date[6:8]}"
        print(f"対象日付: {target_date_formatted}")

        voltage_class = "01"  # 特高・高圧

        # 新しいoneTimeTokenを取得（html2から）
        one_time_token_match2 = re.search(r'id="oneTimeToken" value="([^"]+)"', html2)
        if one_time_token_match2:
            one_time_token2 = one_time_token_match2.group(1)
            print(f"✅ 新しいoneTimeToken: {one_time_token2[:20]}...")
        else:
            print("❌ 新しいトークンが見つかりませんでした")
            return

        # 表示ボタンのリクエスト（AJAXでデータ取得）
        search_url = f"{self.domain}/LNXWPWSS06OH/GPIpG9040/search"

        # JSONペイロードを構築
        import json
        json_payload = {
            "targetDateFrom": target_date_formatted,
            "targetDateTo": target_date_formatted,
            "ppsCd": "",
            "voltageClass": voltage_class,
            "balanceFileType": "",
            "divideFrom": "",
            "divideTo": "",
            "pageNo": "1",
            "sortNo": "0"
        }

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
            "-H", "Content-Type: application/json",
            "-H", f"Referer: {doujidouryou_url}",
            "-H", "Accept: application/json, text/javascript, */*; q=0.01",
            "-H", "Origin: https://nsc.hepco.co.jp",
            "-H", "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36",
            "-H", f"logintoken: {login_token}",
            "-H", f"onetimetoken: {one_time_token2}",
            "-H", "X-Requested-With: XMLHttpRequest",
            "-d", json.dumps(json_payload),
            search_url,
        ]
        result3 = subprocess.run(cmd3, capture_output=True, text=False)

        if result3.returncode != 0:
            print(f"❌ curlエラー: {result3.stderr.decode('utf-8', errors='replace')}")
            return

        # JSONレスポンスを取得
        json_response = result3.stdout.decode('utf-8')
        print("✅ 検索結果取得成功（JSON）")

        # JSONをパース
        try:
            response_data = json.loads(json_response)

        except json.JSONDecodeError as e:
            print(f"❌ JSONパースエラー: {e}")
            return

        # JSONファイルを保存
        out_path3 = os.path.join(self.download_dir, "hokkaido_search_result.json")
        with open(out_path3, "w", encoding="utf-8") as f:
            json.dump(response_data, f, ensure_ascii=False, indent=2)
        print(f"📄 JSONを保存しました: {out_path3}")

        # --- 5️⃣ ファイルをダウンロード ---
        print("\nSTEP5: ZIPファイルをダウンロード")

        # レスポンスからbalanceListを取得
        balance_list = response_data.get("data", {}).get("balanceList", [])
        if not balance_list:
            print("⚠️ ダウンロード対象のファイルが見つかりませんでした")
            return

        print(f"📦 {len(balance_list)} 個のファイルが見つかりました")

        # 新しいoneTimeTokenを取得
        one_time_token3 = response_data.get("oneTimeToken")
        if not one_time_token3:
            print("❌ 新しいトークンが見つかりませんでした")
            return

        # 各ファイルをダウンロード
        for file_info in balance_list:
            balance_id = file_info.get("balanceId")
            file_name = file_info.get("fileName")
            print(f"  📥 {file_name} (ID: {balance_id})")

            # ダウンロードリクエスト
            download_url = f"{self.domain}/LNXWPWSS06OH/GPIpG9040/fileNameDownload"
            download_payload = {"balanceIdList": [str(balance_id)]}

            download_path = os.path.join(self.download_dir, file_name)

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
                "-H", "Content-Type: application/json;charset=UTF-8",
                "-H", f"Referer: {doujidouryou_url}",
                "-H", "Accept: */*",
                "-H", "Origin: https://nsc.hepco.co.jp",
                "-H", "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36",
                "-H", f"logintoken: {login_token}",
                "-H", f"onetimetoken: {one_time_token3}",
                "-H", "X-Requested-With: XMLHttpRequest",
                "-H", "xmlhttprequesttype: downloadFile",
                "-d", json.dumps(download_payload),
                "-o", download_path,
                download_url,
            ]

            result4 = subprocess.run(cmd4, capture_output=True, text=False)

            if result4.returncode == 0:
                if os.path.exists(download_path):
                    file_size = os.path.getsize(download_path)
                    print(f"    ✅ ダウンロード成功: {file_size} bytes")
                else:
                    print(f"    ❌ ファイルが作成されませんでした")
            else:
                print(f"    ❌ ダウンロード失敗: {result4.stderr.decode('utf-8', errors='replace')}")

        print(f"\n✅ ダウンロード完了: {len(balance_list)} ファイル")

        # --- 6️⃣ ZIP解凍 → CSV変換 → S3アップロード ---
        if len(balance_list) > 0:
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
