# scraping.py
'''
スクレイピングの実行はこのファイルで行う
region_scrapingフォルダの中にそれぞれの地方のスクレイピングはまとめてある
'''
from region.kansai_scraping import KansaiScraping
from region.tohoku_scraping import TohokuScraping
from region.hokuriku_scraping import HokurikuScraping
from region.kyusyu_scraping import KyusyuScraping
from region.kanto_scraping import KantoScraping
from region.tyubu_scraping import TyubuScraping
from region.hokkaido_scraping import HokkaidoScraping
from region.shikoku_scraping import ShikokuScraping
from region.tyugoku_scraping import TyugokuScraping

class Scraping:
    def __init__(self, region, config):
        self.region = region
        self.kansai = KansaiScraping(region, config)
        self.tyugoku = TyugokuScraping(region, config)
        self.tohoku = TohokuScraping(region, config)
        self.hokuriku = HokurikuScraping(region, config)
        self.kyusyu = KyusyuScraping(region, config)
        self.hokkaido = HokkaidoScraping(region, config)
        self.shikoku = ShikokuScraping(region, config)
        self.kanto = KantoScraping(region, config)
        self.tyubu = TyubuScraping(region, config)
        
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
            await self.kanto.scraping()
        elif self.region == '中部':
            await self.tyubu.scraping()
        elif self.region == '北海道':
            self.hokkaido.scraping()
        elif self.region == '四国':
            await self.shikoku.scraping()
        
        