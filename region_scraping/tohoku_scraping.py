'''
東北地方のスクレイピングをする
folder_pathにスクレイピングでダウンロードしたファイルを置く
'''
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui
from time import sleep
import threading
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from config import Config
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

class TohokuScraping:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.config = Config()
        self.tohoku_url = self.config.tohoku_url
        self.yesterday_str = self.config.yesterday_str
        self.year2digit = self.config.year2digit
        self.month_day = self.config.month_day
        
    def scraping(self):
        #実際にスクレイピングを行う
        options = Options()

        #options.add_argument("--headless") #EC2にデプロイするときに外す
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")  
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        # download.default_directoryは、EC2のpathに変更
        prefs = {
            "download.default_directory": self.folder_path,  
            "download.prompt_for_download": False,
            "directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)

        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )

        # 証明書許可のために、一定時間後に指定位置をクリックする処理を別スレッドで実行
        def click_cert_ok():
            sleep(2)
            pyautogui.moveTo(585, 378, duration=0.5)
            pyautogui.click()
        threading.Thread(target=click_cert_ok).start()

        # 東北電力のサイトにアクセス
        driver.get(self.tohoku_url)

        # 左フレーム「contents」内で「日毎30分値」リンクをクリック
        WebDriverWait(driver, 40).until(
            EC.frame_to_be_available_and_switch_to_it("contents")
        )
        print("[INFO] 'contents'フレームに切り替えました。")
        
        links = driver.find_elements(By.TAG_NAME, "a")
        for link in links:
            if "日毎30分値" in link.text:
                link.click()
                break

        sleep(15)

        # メインフレームへ戻り、ファイル一覧を処理
        driver.switch_to.default_content()
        WebDriverWait(driver, 30).until(
            EC.frame_to_be_available_and_switch_to_it("main")
        )
        # 対象ZIPファイル名の生成（例: "W401202025041500000000.zip"）
        target_filename = f"W4012020{self.year2digit}{self.month_day}00000000.zip"
        print(f"[DEBUG] 探索対象ファイル名: {target_filename}")

        # JavaScriptの__doPostBackイベントが設定されているリンクから、対象ファイルのダウンロードリンクを検索
        zip_links = driver.find_elements(By.XPATH, '//a[starts-with(@href,"javascript:__doPostBack")]')
        found = False
        for link in zip_links:
            if target_filename in link.text:
                event_target = link.get_attribute("href").split("'")[1]
                driver.execute_script(f"__doPostBack('{event_target}', '')")
                print(f"[INFO] ダウンロードをトリガーしました: {target_filename}")
                found = True
                break

        if not found:
            print(f"[WARNING] 対象ファイルが見つかりませんでした: {target_filename}")

        sleep(20)
        driver.quit()