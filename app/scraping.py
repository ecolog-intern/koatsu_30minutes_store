'''
スクレイピングの実行はこのファイルで行う
region_scrapingフォルダの中にそれぞれの地方のスクレイピングはまとめてある
'''
from region_scraping.kansai_scraping import KansaiScraping
from region_scraping.tohoku_scraping import TohokuScraping
from region_scraping.tyugoku_scraping import TyugokuScraping
from region_scraping.hokuriku_scraping import HokurikuScraping
from region_scraping.kyusyu_scraping import KyusyuScraping
from region_scraping.tokyo_scraping import TokyoScraping
from region_scraping.tyubu_scraping import TyubuScraping


class Scraping:
    def __init__(self, region, folder_path):
        self.region = region
        self.folder_path = folder_path
        self.kansai = KansaiScraping(folder_path)
        self.tohoku = TohokuScraping(folder_path)
        self.tyugoku = TyugokuScraping(folder_path)
        self.hokuriku = HokurikuScraping(folder_path)
        self.kyusyu = KyusyuScraping(folder_path)
        self.tokyo = TokyoScraping(folder_path)
        self.tyubu = TyubuScraping(folder_path)
        
    def scraping(self):
        #実際にはここでregionごとにスクレイピングを実行する
        if self.region == '中国':
            self.tyugoku.scraping()
        elif self.region == '関西':
            self.kansai.scraping()
        elif self.region == '東北':
            self.tohoku.scraping()
        elif self.region == '北陸':
            self.hokuriku.scraping()
        elif self.region == '九州':
            self.kyusyu.scraping()
        elif self.region == '関東':
            self.tokyo.scraping()
        elif self.region == '中部':
            self.tyubu.scraping()
        
        