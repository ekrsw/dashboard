import pandas as pd
import datetime
import logging
import settings
from typing import List, Callable, Any, Optional
import asyncio
from functools import partial, wraps
from concurrent.futures import ThreadPoolExecutor
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import (
    WebDriverException,
    NoSuchElementException,
    TimeoutException
)

# カスタム例外の定義
class ScraperError(Exception):
    """スクレイピングに関するカスタム例外"""
    pass

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)  # 必要に応じて設定

def async_retry(max_attempts: int = 3, delay: float = 1.0, exceptions: tuple = (Exception,)):
    """
    非同期関数用のリトライデコレーター
    """
    def decorator(func: Callable):
        @wraps(func)
        async def wrapper(*args, **kwargs):
            attempt = 0
            while attempt < max_attempts:
                try:
                    return await func(*args, **kwargs)
                except exceptions as e:
                    attempt += 1
                    logger.warning(f"Attempt {attempt} failed for {func.__name__}: {e}")
                    if attempt >= max_attempts:
                        logger.error(f"All {max_attempts} attempts failed for {func.__name__}")
                        raise ScraperError(f"Operation '{func.__name__}' failed after {max_attempts} attempts") from e
                    await asyncio.sleep(delay)
        return wrapper
    return decorator

class BaseScraper:
    def __init__(self, url: str, id: str, executor: Optional[ThreadPoolExecutor] = None):
        self.url = url
        self.id = id
        self.df = pd.DataFrame()
        self.driver = None
        self.loop = asyncio.get_event_loop()
        self.executor = executor or ThreadPoolExecutor(max_workers=5)  # 適宜調整

    async def __aenter__(self):
        await self.create_driver()
        return self

    async def __aexit__(self, exc_type, exc, tb):
        await self.close_driver()

    @async_retry(max_attempts=3, delay=2.0, exceptions=(ScraperError, WebDriverException))
    async def fetch_data(self):
        try:
            await self.login()
            # 必要に応じて他のスクレイピング手順を非同期的に実行
            # 例:
            # await self.call_template(['テンプレート1', '値1'])
            # await self.filter_by_date(datetime.date.today() - datetime.timedelta(days=7), datetime.date.today())
            # データの取得と処理
        except ScraperError as e:
            logger.error(f"データの取得に失敗しました: {e}")
            raise
        except Exception as e:
            logger.error(f"予期しないエラーが発生しました: {e}")
            raise ScraperError("予期しないエラーが発生しました") from e

    @async_retry(max_attempts=3, delay=2.0, exceptions=(WebDriverException,))
    async def create_driver(self):
        try:
            if self.driver:
                await self.close_driver()

            options = Options()

            # ブラウザを表示させない。
            if settings.HEADLESS_MODE:
                options.add_argument('--headless')

            # コマンドプロンプトのログを表示させない。
            options.add_argument('--disable-logging')
            options.add_argument('--disable-extensions')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--log-level=3')
            options.add_experimental_option('excludeSwitches', ['enable-logging'])

            driver_creation = partial(webdriver.Chrome, options=options)
            self.driver = await self.loop.run_in_executor(self.executor, driver_creation)
            self.driver.implicitly_wait(5)
            logger.info("Webドライバーを正常に作成しました。")
        except Exception as e:
            logger.error(f"ドライバーの作成に失敗しました。: {e}")
            raise ScraperError("ドライバーの作成に失敗しました") from e

    @async_retry(max_attempts=3, delay=2.0, exceptions=(NoSuchElementException, ScraperError))
    async def login(self):
        """レポータに接続してログイン"""
        try:
            await self.loop.run_in_executor(self.executor, self.driver.get, self.url)
            logger.info(f"URL {self.url} にアクセスしました。")

            # ID入力
            logon_operator_id = await self.find_element(By.ID, 'logon-operator-id')
            await self.loop.run_in_executor(self.executor, logon_operator_id.send_keys, self.id)
            logger.info("IDを入力しました。")

            # ログインボタンをクリック
            logon_btn = await self.find_element(By.ID, 'logon-btn')
            await self.loop.run_in_executor(self.executor, logon_btn.click)
            logger.info("ログインボタンをクリックしました。")

            # 必要に応じてログイン成功の確認ステップを追加
            await asyncio.sleep(2)  # ログイン処理の待機（適宜調整）
        except Exception as e:
            logger.error(f"ログインに失敗しました。: {e}")
            raise ScraperError("ログインに失敗しました") from e

    @async_retry(max_attempts=3, delay=2.0, exceptions=(NoSuchElementException, ScraperError))
    async def call_template(self, template: List[str]) -> None:
        """
        テンプレート呼び出し、指定の集計期間を表示
        """
        try:
            # テンプレートタイトルをクリック
            template_title = await self.find_element(By.ID, 'template-title-span')
            await self.loop.run_in_executor(self.executor, template_title.click)
            logger.info("テンプレートタイトルをクリックしました。")

            # ダウンロード範囲セレクト
            el1 = await self.find_element(By.ID, 'download-open-range-select')
            s1 = Select(el1)
            await self.loop.run_in_executor(self.executor, partial(s1.select_by_visible_text, template[0]))
            logger.info(f"ダウンロード範囲を '{template[0]}' に設定しました。")

            # テンプレートダウンロードセレクト
            el2 = await self.find_element(By.ID, 'template-download-select')
            s2 = Select(el2)
            await self.loop.run_in_executor(self.executor, partial(s2.select_by_value, template[1]))
            logger.info(f"テンプレートダウンロードを '{template[1]}' に設定しました。")

            # テンプレート作成ボタンをクリック
            template_creation_btn = await self.find_element(By.ID, 'template-creation-btn')
            await self.loop.run_in_executor(self.executor, template_creation_btn.click)
            logger.info("テンプレート作成ボタンをクリックしました。")

            # 必要に応じて処理の完了を待機
            await asyncio.sleep(2)
        except Exception as e:
            logger.error(f"テンプレート呼び出しに失敗しました。: {e}")
            raise ScraperError("テンプレート呼び出しに失敗しました") from e

    @async_retry(max_attempts=3, delay=2.0, exceptions=(NoSuchElementException, ScraperError))
    async def filter_by_date(self, start_date: datetime.date,
                             end_date: datetime.date,
                             input_id: str = "0") -> None:
        try:
            from_id = f'panel-td-input-from-date-{input_id}'
            to_id = f'panel-td-input-to-date-{input_id}'
            create_report_id = f'panel-td-create-report-{input_id}'

            # 集計期間のfromをクリアしてfrom_dateを送信
            from_input = await self.find_element(By.ID, from_id)
            await self.loop.run_in_executor(self.executor, partial(from_input.send_keys, Keys.CONTROL + 'a'))
            await self.loop.run_in_executor(self.executor, partial(from_input.send_keys, Keys.DELETE))
            await self.loop.run_in_executor(self.executor, partial(from_input.send_keys, start_date.strftime('%Y/%m/%d')))
            logger.info(f"開始日を {start_date} に設定しました。")

            # 集計期間のtoをクリアしてto_dateを送信
            to_input = await self.find_element(By.ID, to_id)
            await self.loop.run_in_executor(self.executor, partial(to_input.send_keys, Keys.CONTROL + 'a'))
            await self.loop.run_in_executor(self.executor, partial(to_input.send_keys, Keys.DELETE))
            await self.loop.run_in_executor(self.executor, partial(to_input.send_keys, end_date.strftime('%Y/%m/%d')))
            logger.info(f"終了日を {end_date} に設定しました。")

            # レポート作成ボタンをクリック
            create_report_button = await self.find_element(By.ID, create_report_id)
            await self.loop.run_in_executor(self.executor, create_report_button.click)
            logger.info("レポート作成ボタンをクリックしました。")

            # 必要に応じて処理の完了を待機
            await asyncio.sleep(2)
        except Exception as e:
            logger.error(f"日付フィルタリングに失敗しました。: {e}")
            raise ScraperError("日付フィルタリングに失敗しました") from e

    @async_retry(max_attempts=3, delay=2.0, exceptions=(NoSuchElementException, ScraperError))
    async def select_tabs(self, tab_id_num: str = "2"):
        """
        タブを選択するメソッド
        """
        try:
            tab_id = f'normal-title{tab_id_num}'
            tab_element = await self.find_element(By.ID, tab_id)
            await self.loop.run_in_executor(self.executor, tab_element.click)
            logger.info(f"タブ '{tab_id}' を選択しました。")

            # 必要に応じて処理の完了を待機
            await asyncio.sleep(1)
        except Exception as e:
            logger.error(f"タブ選択に失敗しました。: {e}")
            raise ScraperError("タブ選択に失敗しました") from e

    async def find_element(self, by: By, value: str, timeout: int = 10):
        """
        要素を検索し、見つからない場合はタイムアウトを待つ。
        """
        for attempt in range(timeout):
            try:
                element = await self.loop.run_in_executor(self.executor, partial(self.driver.find_element, by, value))
                if element:
                    return element
            except NoSuchElementException:
                await asyncio.sleep(1)  # 再試行前に待機
        logger.error(f"要素が見つかりません: {by}={value}")
        raise ScraperError(f"要素が見つかりません: {by}={value}")

    async def close_driver(self):
        """ドライバーを閉じる"""
        if self.driver:
            try:
                await self.loop.run_in_executor(self.executor, self.driver.quit)
                logger.info("ドライバーを正常に閉じました。")
            except Exception as e:
                logger.error(f"ドライバーの閉鎖に失敗しました: {e}")
            finally:
                self.driver = None

    async def quit_driver(self):
        """ドライバーをクイットする補助メソッド"""
        await self.close_driver()
