'''
関西地方のスクレイピングをする
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

class KansaiScraping:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.config = Config()
        self.kansai_url = self.config.kansai_url
        self.kansai_error_url = self.config.kansai_error_url
        
    def scraping(self):
        #実際にスクレイピングを行う
        options = webdriver.ChromeOptions()

        # ★ プロキシ設定
        options.add_argument('--proxy-server=http://proxy:8080')
        options.add_argument('--disable-gpu')
        options.add_argument('--ignore-certificate-errors')

        # ★ ダウンロード先指定（prefsの設定）
        prefs = {
            "download.default_directory": self.folder_path,
            "download.prompt_for_download": False,  # ダウンロード時の確認ダイアログを無効化
            "directory_upgrade": True
        }
        options.add_experimental_option("prefs", prefs)

        driver = webdriver.Chrome(options=options)
        
        # 証明書の許可処理: 指定した位置にマウスを移動し、クリックを実施
        pyautogui.moveTo(585, 378, duration=0.5)
        def click_cert_ok():
            sleep(2)
            pyautogui.moveTo(585, 378, duration=0.5)
            pyautogui.click()
        
        click_thread = threading.Thread(target=click_cert_ok)
        click_thread.start()
        
        # 関西電力のサイトへアクセス
        driver.get(self.kansai_url)
        sleep(10)
        
        # 関西電力のホームページ操作:
        # 1. "wideBtn4" クラスの最初のボタンをクリック
        elem_button = driver.find_elements(By.CLASS_NAME, "wideBtn4")[0]
        elem_button.click()
        sleep(1)
        
        # 2. エラー画面の場合、"smallBtn6" クラスのボタンをクリック
        if driver.current_url == self.kansai_error_url:
            elem_button = driver.find_elements(By.CLASS_NAME, "smallBtn6")[0]
            elem_button.click()
            sleep(2)
        
        # 3. 「同時同量支援」ボタンのクリック
        elem_doji = driver.find_element(By.NAME, "DojiDoryouShien")
        elem_doji.click()
        sleep(2)
        
        # 4. テーブル内からZIPファイルのリンクを特定してクリック（ZIPファイルのダウンロード）
        elem_body = driver.find_elements(By.ID, 'BODY1')[1]
        elem_body = elem_body.find_element(By.TAG_NAME, 'tbody')
        elem_file = elem_body.find_elements(By.TAG_NAME, 'tr')[1]
        elem_file = elem_file.find_element(By.TAG_NAME, 'a')  # ZIPファイルのダウンロードリンク
        elem_file.click()
        
        sleep(10)
        driver.close() 