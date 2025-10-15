# main.py
import os
from config import Config
import asyncio
from scraping import Scraping

async def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    config = Config()

    all_area = ['北海道'] # '中国','東北','北陸','九州','中部','関東'
    # all_area = ['関西','四国']
    for area in all_area:
        print(f"========{area}の処理中========")
        scraping = Scraping(area, config)
        await scraping.scraping()
        print(f"=============================")

if __name__ == '__main__':
    asyncio.run(main())
    