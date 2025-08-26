from datetime import datetime, timedelta

class Config:
    def __init__(self):
        '''
        フォルダパスの一覧
        '''
        #顧客一覧
        self.consumer_filtering_file = 'consumer_region_filtering.xlsx'
        #self.consumer_filtering_file = "filtering_test.xlsx"
        #作成したファイルの掃き出し先(クロームのスクレイピングの関係で相対パスだときついかも)
        self.base_folder = r'\\f-hikari02\法人事業本部\事業開発部\02.アライアンス事業部\31.ガスでん\27.インターン用\20.高圧30分値\高圧実行\フォルダ保存用'
        #エラーログ
        self.error_log = "error_log.txt"
        
        '''
        電力会社url
        '''
        self.tyugoku_url = 'https://takusouhp.energia.co.jp/COMM/xhtml/COMMLOP.xhtml'
        self.tohoku_url = 'https://takuso2-web.takuso.tohoku-epco.co.jp/G83_PPS/'
        self.kansai_url = 'https://www4.kepco.co.jp/'
        #関西はなぜかエラーが出るのでエラー処理用のurlを用意
        self.kansai_error_url = 'https://www4.kepco.co.jp/takusouinfo/H24DF700A04.do'
        self.tokyo_url = 'https://pu00.www6.tepco.co.jp/org_web/LVA2RG/pgsslogin.faces'
        self.tyubu_url = 'https://epcdss-www.chuden.co.jp/46264/'
        self.hokuriku_url = 'https://wsweb4.rikuden.co.jp/tfx/tfxo110/tfxo110s010'
        self.kyusyu_url = 'https://nsc-www.network.kyuden.co.jp/BP_WEB_SERVER/'
        
        
        '''
        日付系
        '''
        today = datetime.today()
        yesterday = today - timedelta(days=1)
        self.yesterday_str = yesterday.strftime("%Y%m%d")
        self.yesterday_month_str = yesterday.strftime("%Y%m")
        self.month_day = yesterday.strftime('%m%d')  
        self.year2digit = yesterday.strftime('%y') 
        self.yesterday_tsuki_nichi = f"{yesterday.month}月{yesterday.day}日"
        
        '''
        メール系
        '''
        self.mailaddress = ["kousuke_morita@eco-log.co.jp"]