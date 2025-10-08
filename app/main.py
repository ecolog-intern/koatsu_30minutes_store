import os
import asyncio
from datetime import datetime, timedelta
from scraping import Scraping

async def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    #日付の設定
    today = datetime.today()
    yesterday = today - timedelta(days=1)
    yesterday_str = yesterday.strftime("%Y%m%d")
    yesterday_month_str = yesterday.strftime("%Y%m")
    month_day = yesterday.strftime('%m%d')  
    year2digit = yesterday.strftime('%y') 
    yesterday_tsuki_nichi = f"{yesterday.month}月{yesterday.day}日"

    all_area = ['関西']
    # all_area = ['中国','関西','東北','北陸','九州','関東','中部', '北海道', '四国']

    for area in all_area:
        scraping = Scraping(area, yesterday_month_str, yesterday_str)
        await scraping.scraping()


if __name__ == '__main__':
    asyncio.run(main())
    