'''
スクレイピングの実行はこのファイルで行う
region_scrapingフォルダの中にそれぞれの地方のスクレイピングはまとめてある
'''
from region.kansai_scraping import KansaiScraping
# from region.tohoku_scraping import TohokuScraping
# from region.hokuriku_scraping import HokurikuScraping
# from region.kyusyu_scraping import KyusyuScraping
# from region.tokyo_scraping import TokyoScraping
# from region.tyubu_scraping import TyubuScraping

# from region.tyugoku_scraping import TyugokuScraping

class Scraping:
    def __init__(self, region, yesterday_month_str, yesterday_str):
        self.region = region
        self.yesterday_month_str = yesterday_month_str
        self.yesterday_str = yesterday_str
        self.kansai = KansaiScraping(self.region, self.yesterday_month_str, self.yesterday_str)
        # self.tohoku = TohokuScraping(self.region, self.yesterday_month_str, self.yesterday_str)
        # self.hokuriku = HokurikuScraping(self.region, self.yesterday_month_str, self.yesterday_str)
        # self.kyusyu = KyusyuScraping(self.region, self.yesterday_month_str, self.yesterday_str)
        # self.tokyo = TokyoScraping(self.region, self.yesterday_month_str, self.yesterday_str)
        # self.tyubu = TyubuScraping(self.region, self.yesterday_month_str, self.yesterday_str)
        # self.tyugoku = TyugokuScraping(self.region, self.yesterday_month_str, self.yesterday_str)
        
    async def scraping(self):
        #実際にはここでregionごとにスクレイピングを実行する
        if self.region == '中国':
            await self.tyugoku.scraping()
        elif self.region == '関西':
            self.kansai.scraping()
        elif self.region == '東北':
            await self.tohoku.scraping()
        elif self.region == '北陸':
            await self.hokuriku.scraping()
        elif self.region == '九州':
            await self.kyusyu.scraping()
        elif self.region == '関東':
            await self.tokyo.scraping()
        elif self.region == '中部':
            await self.tyubu.scraping()
        
        