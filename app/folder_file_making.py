#フォルダを作成して、エクセルまで作成
import os
from excel_making import MonthExcelMaking
from config import Config

class FolderFileMaking:
    def __init__(self, base_folder, index):
        self.config =Config()
        
        self.base_folder = base_folder
        self.index = index
        self.yesterday_month_str = self.config.yesterday_month_str
        self.yesterday_str = self.config.yesterday_str

        
    def folder_making(self):
        #フォルダーを作成して、ファイルのパスも作成する
        #外部からもこの関数で呼び出す
        
        #基準となるフォルダを作成
        index_folder = os.path.join(self.base_folder, self.index)
        if not os.path.exists(index_folder):
            os.makedirs(index_folder)
        
        #当月用のフォルダを作成（例: base_folder/index/YYYYMM）
        month_folder = os.path.join(index_folder, self.yesterday_month_str)
        if not os.path.exists(month_folder):
            os.makedirs(month_folder)
        
        '''
        #Excelファイルの作成または存在チェック
        # ファイル名例: index_YYYYMM.xlsx
        excel_name = f'{self.index}_{self.yesterday_month_str}.xlsx'
        excel_path = os.path.join(month_folder, excel_name)
        excel_making = MonthExcelMaking(excel_path)
        if not os.path.exists(excel_path):
            excel_making.excel_making()
        '''
        
        #昨日の日付用のフォルダを作成（例: base_folder/index/YYYYMM/YYYYMMDD）
        yesterday_folder_path = os.path.join(index_folder, self.yesterday_month_str, self.yesterday_str)
        if not os.path.exists(yesterday_folder_path):
            os.makedirs(yesterday_folder_path)
        
        # 昨日フォルダのパスとExcelファイルのパスを返す
        return yesterday_folder_path