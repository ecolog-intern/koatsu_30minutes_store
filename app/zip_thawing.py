# zip_thawing.py
"""
指定されたフォルダ内のZIPファイルを検索し、最新のZIPファイルを解凍して、
解凍されたXMLファイルをCSVに変換し、同じフォルダに保存する。
"""
import os
import zipfile
from time import sleep
from config import Config
from xml_to_csv import xml_to_csv

class ZipThawing:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.config = Config()

    def zip_thawing(self):
        """
        最新のZIPファイルを解凍し、XML→CSV変換して同じフォルダに保存。
        """
        extract_to = self.folder_path
        extracted_file = None

        # --- ZIPファイルの候補を取得 ---
        files = [os.path.join(self.folder_path, f) for f in os.listdir(self.folder_path)]
        zip_files = [f for f in files if f.lower().endswith(".zip")]

        if not zip_files:
            raise FileNotFoundError(f"[ERROR] {self.folder_path} にZIPファイルが存在しません。")

        # --- 作成日時の降順で並べ替え ---
        zip_files.sort(key=os.path.getctime, reverse=True)

        # --- 最新ZIPを解凍 ---
        for zip_path in zip_files:
            try:
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_to)
                    namelist = zip_ref.namelist()
                    if not namelist:
                        continue
                    # 最初に見つかったXMLを採用
                    extracted_name = namelist[0]
                    extracted_file = os.path.join(extract_to, extracted_name)
                    print(f"✅ 解凍成功: {extracted_file}")
                    break
            except zipfile.BadZipFile:
                print(f"⚠️ 無効なZIPファイルをスキップ: {zip_path}")
                continue

        # --- 解凍ファイルが見つからない場合 ---
        if not extracted_file or not os.path.exists(extracted_file):
            raise FileNotFoundError("[ERROR] ZIPファイル内に有効なファイルが見つかりませんでした。")
        sleep(2)

        # xml_to_csv は「変換後パス」を返す想定
        csv_path = xml_to_csv(extracted_file, extract_to)

        # --- 保存確認 ---
        if os.path.exists(csv_path):
            print(f"✅ CSV変換完了: {csv_path}")
        else:
            print(f"⚠️ CSVファイルが見つかりません: {csv_path}")

        return csv_path
