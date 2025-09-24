'''
指定されたフォルダ内のZIPファイルを検索し、最新のZIPファイルを解凍して、
解凍されたファイルをExcel形式に変換・保存します。
'''
import os
import zipfile
from time import sleep
import win32com.client
from config import Config
from xml_to_csv import xml_to_csv

class ZipThawing:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.config = Config()
        self.yesterday_str = self.config.yesterday_str
        
    def zip_thawing(self):
        #zipファイルはこの関数で解凍する
        extract_to = self.folder_path
        extracted_file = None

        # フォルダ内のファイルをフルパスで取得し、作成日時の降順でソート
        files = [os.path.join(self.folder_path, f) for f in os.listdir(self.folder_path)]
        files = sorted(files, key=os.path.getctime, reverse=True)

        for file_item in files:
            print(f"  - {file_item}")

        # 各ファイルについて、ZIPファイルとして解凍を試みる
        for file_path in files:
            try:
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    # ZIPファイル内のすべてのファイルを展開先に解凍
                    zip_ref.extractall(extract_to)
                    namelist = zip_ref.namelist()
                    if not namelist:
                        continue
                    # 最初に見つかったファイルを採用する
                    extracted_name = namelist[0]
                    extracted_file = os.path.join(extract_to, extracted_name)
                    break
            except zipfile.BadZipFile:
                continue

        # 有効なファイルが見つからなければエラーを発生させる
        if not extracted_file:
            raise FileNotFoundError("[ERROR] ZIPファイルの中に有効なファイルが見つかりませんでした。")

        sleep(2)

        csv_path = xml_to_csv(extracted_file, self.folder_path)
        
        return csv_path