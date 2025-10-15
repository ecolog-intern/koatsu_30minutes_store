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
