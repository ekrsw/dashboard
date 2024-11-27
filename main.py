from src.excel_sync import SynchronizedExcelProcessor
from src.processors.activity_processor import ActivityProcessor
from src.processors.support_processor import SupportProcessor
from src.scrapers.base_scraper import BaseScraper
import settings
import asyncio
import datetime

import logging

LOG_FILE = settings.LOG_FILE

def setup_logging(log_file):
    # ロギングの設定
    logging.basicConfig(
        level=logging.DEBUG,  # ログレベルをDEBUGに設定
        format='%(asctime)s:%(levelname)s:%(name)s:%(message)s',
        handlers=[
            logging.FileHandler(log_file, mode='a', encoding='utf-8'),  # ファイルへのログ出力
            logging.StreamHandler()  # コンソールへのログ出力
        ]
    )

setup_logging(LOG_FILE)
logger = logging.getLogger(__name__)
activity_file_path = settings.ACTIVITY_FILE
support_file_path = settings.SUPPORT_FILE
close_file_path = settings.CLOSE_FILE

async def my_task():
    scraper = BaseScraper(settings.REPORTER_URL, settings.REPORTER_ID)

    try:
        await scraper.create_driver()
        await scraper.login()
        await scraper.call_template(template=settings.TEMPLATE_TVS)
        await scraper.filter_by_date(start_date=datetime.date.today(),
                                     end_date=datetime.date.today(),
                                     input_id=0)
        await asyncio.sleep(5)
        await scraper.select_tabs(tab_id_num="2")
        await scraper.filter_by_date(start_date=datetime.date.today(),
                                     end_date=datetime.date.today(),
                                     input_id=1)
        await asyncio.sleep(5)
    
    except Exception as e:
        print(f"エラーが発生しました: {e}")
    
    finally:
        await scraper.close_driver()

async def main():
    # 同期するExcelファイルのリスト
    excel_files = [
        activity_file_path,
        support_file_path,
        close_file_path
    ]

    # Excel同期プロセッサのインスタンスを作成
    excel_processor = SynchronizedExcelProcessor(excel_files)
    excel_processor.start()

    # 並行して実行するタスク
    await asyncio.gather(
        my_task(),
        # 必要に応じて他の非同期タスクを追加
    )

    # Excel同期処理の完了を待つ
    while excel_processor.thread.is_alive():
        logger.info("Excel同期処理が完了するのを待っています...")
        await asyncio.sleep(1)

    logger.info("全てのタスクが完了しました。")


    ap = ActivityProcessor(settings.ACTIVITY_FILE)
    ap.load_data()
    ap.process()

    sp = SupportProcessor(settings.SUPPORT_FILE)
    sp.load_data()
    sp.process()

    print('TVS 20以内', ap.cb_0_20_tvs)
    print('TVS 直受け', sp.direct_tvs)
    print('TVS 20～30',ap.cb_20_30_tvs)

# 非同期イベントループの実行
if __name__ == "__main__":
    asyncio.run(main())
