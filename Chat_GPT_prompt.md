## ディレクトリ構成
```
dashboard/
├── data/
│   ├── TS_todays_activity.xlsx
│   ├── TS_todays_close.xlsx
│   └── TS_todays_support.xlsx
├── src/
│   ├── __init__.py
│   ├── excel_sync.py
│   ├── web_scraper.py
│   ├── kpi_calculator.py
│   ├── output_handler.py
│   └── processors/
│       ├── __init__.py
│       ├── activity_processor.py
│       ├── close_processor.py
│       └── support_processor.py
├── tests/
│   ├── __init__.py
│   ├── test_excel_sync.py
│   └── test_output_handler.py
├── main.py
├── requirements.txt
├── settings.py
└── README.md
```
## src/processors/activity_processor.py
```src/processors/activity_processor.py
from .base_processor import BaseProcessor
import pandas as pd
import datetime
import settings
import logging

logger = logging.getLogger(__name__)

class ActivityProcessor(BaseProcessor):
    def process(self):
        """
        Activityファイルのデータを指定された条件でフィルタリングおよび整形します。
        """
        try:
            # '件名'列を文字列型に変換
            self.df['件名'] = self.df['件名'].astype(str)

            # 件名に「【受付】」が含まれていないもののみ残す
            self.df = self.df[~self.df['件名'].str.contains('【受付】', na=False)]
            logger.info(f"Filtered out rows containing '【受付】'. Remaining rows: {self.df.shape[0]}")

            # 日付範囲でフィルタリング
            start_date = datetime.date.today()
            end_date = datetime.date.today()
            self.df = self.filtered_by_date_range(self.df, start_date, end_date)
            logger.info(f"Filtered DataFrame by date range {start_date} to {end_date}. Remaining rows: {self.df.shape[0]}")

            # 案件番号でソートし、最も早い日時を残して重複を削除
            self.df = self.df.sort_values(by=['案件番号 (関連) (サポート案件)', '登録日時'])
            self.df = self.df.drop_duplicates(subset='案件番号 (関連) (サポート案件)', keep='first')
            logger.info(f"Sorted and dropped duplicates. Remaining rows: {self.df.shape[0]}")

            # '登録日時'と'登録日時 (関連) (サポート案件)'をdatetime型に変換
            self.df['登録日時'] = pd.to_datetime(self.df['登録日時'], unit='d', origin='1899-12-30')
            self.df['登録日時 (関連) (サポート案件)'] = pd.to_datetime(self.df['登録日時 (関連) (サポート案件)'], unit='d', origin='1899-12-30')

            # 時間差を計算
            self.df['時間差'] = self.df['登録日時'] - self.df['登録日時 (関連) (サポート案件)']
            self.df['時間差'] = self.df['時間差'].fillna(pd.Timedelta(seconds=0))
            logger.info("Calculated '時間差' column.")

            # 必要に応じて他の整形処理を追加

        except Exception as e:
            logger.error(f"Error processing activity data: {e}")
            raise

    def filtered_by_date_range(self, df: pd.DataFrame, start_date: datetime.date, end_date: datetime.date) -> pd.DataFrame:
        """
        start_dateからend_dateの範囲のデータを抽出

        :param df: フィルタリング対象のDataFrame
        :param start_date: 抽出開始日
        :param end_date: 抽出終了日
        :return: フィルタリング後のDataFrame
        """
        start_date_serial = self.datetime_to_serial(datetime.datetime.combine(start_date, datetime.time.min))
        end_date_serial = self.datetime_to_serial(datetime.datetime.combine(end_date + datetime.timedelta(days=1), datetime.time.min))
        
        filtered_df = df[
            (df['登録日時 (関連) (サポート案件)'] >= start_date_serial) &
            (df['登録日時 (関連) (サポート案件)'] < end_date_serial)
        ].reset_index(drop=True)
        
        logger.debug(f"Filtered DataFrame from {start_date} to {end_date}: {filtered_df.shape[0]} rows")
        return filtered_df

    @staticmethod
    def datetime_to_serial(dt: datetime.datetime, base_date=datetime.datetime(1899, 12, 30)) -> float:
        """
        datetimeオブジェクトをシリアル値に変換する。

        :param dt: 変換するdatetimeオブジェクト
        :param base_date: シリアル値の基準日（デフォルトは1899年12月30日）
        :return: シリアル値
        """
        return (dt - base_date).total_seconds() / (24 * 60 * 60)
```

## src/excel_sync.py
```src/excel_sync.py
import win32com.client
import os
from typing import List
import logging
import time
import threading
import pythoncom  # COM初期化に必要
import settings  # settings.py をインポート

# ロガーの設定
logger = logging.getLogger(__name__)

class SynchronizedExcelProcessor:
    def __init__(self, file_paths: List[str], max_retries: int = settings.SYNC_MAX_RETRIES,
                 retry_delay: float = settings.SYNC_RETRY_DELAY,
                 refresh_interval: int = settings.REFRESH_INTERVAL):
        """
        Excelファイルの同期処理を管理するクラス。

        Parameters
        ----------
        file_paths : List[str]
            同期するExcelファイルのパスのリスト。
        max_retries : int, optional
            同期失敗時の最大リトライ回数（デフォルトは設定ファイルから）。
        retry_delay : float, optional
            リトライ間の待機時間（秒、デフォルトは設定ファイルから）。
        refresh_interval : int, optional
            CalculationState を確認する際の待機時間（秒、デフォルトは設定ファイルから）。
        """
        self.file_paths = file_paths
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.refresh_interval = refresh_interval
        self.thread = None
        self.stop_event = threading.Event()

    def start(self) -> None:
        """
        同期処理を別スレッドで開始します。

        Starts the synchronization process in a separate thread.
        """
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()
        logger.info("Excel同期処理スレッドを開始しました。")

    def _run(self):
        """
        同期処理を実行する内部メソッド。

        Manages the synchronization of Excel files, handling retries and exceptions.
        """
        try:
            # COMライブラリを初期化
            pythoncom.CoInitialize()
            logger.debug("COMライブラリを初期化しました。")

            # Excelアプリケーションを作成
            excel = self._create_excel_app()

            # 各ファイルパスに対して処理を実行
            for file_path in self.file_paths:
                # 停止イベントがセットされているか確認
                if self.stop_event.is_set():
                    logger.info("同期処理が停止されました。")
                    break

                # ファイルが存在するか確認
                if not os.path.exists(file_path):
                    logger.warning(f"ファイルが存在しません: {file_path}")
                    continue

                logger.info(f"{file_path} の同期を開始します。")
                retries = 0

                while retries < self.max_retries:
                    try:
                        # ワークブックを開く
                        workbook = excel.Workbooks.Open(file_path)
                        
                        # データの更新を実行
                        logger.debug("Workbook.RefreshAll() を実行します。")
                        workbook.RefreshAll()

                        # 更新が完了するまで待機
                        time.sleep(settings.REFRESH_INTERVAL)

                        # ワークブックを保存して閉じる
                        workbook.Save()
                        workbook.Close()
                        logger.info(f"{file_path} の同期が完了しました。")
                        break  # 成功したのでリトライループを抜ける

                    except Exception as e:
                        retries += 1
                        logger.error(f"{file_path} の同期中にエラーが発生しました（{retries} 回目）: {e}")

                        if retries >= self.max_retries:
                            logger.error(f"{file_path} の同期に{self.max_retries}回失敗しました。Excelを終了します。")
                            try:
                                excel.Quit()
                                logger.info("Excelアプリケーションを終了しました。")
                            except Exception as quit_e:
                                logger.warning(f"Excelの終了中にエラーが発生しました: {quit_e}")

                            # 次のファイルの処理のためにExcelを再起動
                            excel = self._create_excel_app()
                        else:
                            logger.info(f"{file_path} の同期を再試行します。")
                            time.sleep(self.retry_delay)  # リトライ前に待機

            # 全てのファイル処理が完了した後、Excelを終了
            try:
                excel.Quit()
                logger.info("Excelアプリケーションを終了します。")
            except Exception as e:
                logger.warning(f"Excelの終了中にエラーが発生しました: {e}")

        except Exception as e:
            logger.error(f"同期処理中に予期しないエラーが発生しました: {e}")
        finally:
            # COMライブラリを終了
            pythoncom.CoUninitialize()
            logger.debug("COMライブラリを終了しました。")

    def _create_excel_app(self):
        """
        Excelアプリケーションを起動し、設定を行うヘルパーメソッド。

        Returns
        -------
        excel_app : COMObject
            起動したExcelアプリケーションオブジェクト。
        """
        logger.info("Excelアプリケーションを起動します。")
        excel_app = win32com.client.DispatchEx("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        return excel_app

    def stop(self):
        """
        同期処理を停止します。

        Stops the synchronization process.
        """
        self.stop_event.set()
        if self.thread and self.thread.is_alive():
            self.thread.join()
            logger.info("Excel同期処理スレッドを停止しました。")
```
## tests/test_excel_sync.py
```tests/test_excel_sync.py
import unittest
from unittest.mock import MagicMock, patch
from src.excel_sync import SynchronizedExcelProcessor
import threading
import settings

class TestSynchronizedExcelProcessor(unittest.TestCase):
    def setUp(self):
        """
        テストのセットアップを行います。
        """
        self.file_paths = [settings.ACTIVITY_FILE, settings.SUPPORT_FILE, settings.CLOSE_FILE]
        self.processor = SynchronizedExcelProcessor(self.file_paths, max_retries=2, retry_delay=0.1)

    @patch('threading.Thread')
    def test_initialization(self, mock_thread):
        """
        SynchronizedExcelProcessorの初期化をテストします。

        Parameters
        ----------
        mock_thread : MagicMock
            threading.Threadのモックオブジェクト。
        """
        # SynchronizedExcelProcessorのインスタンスを作成
        processor = SynchronizedExcelProcessor(self.file_paths, max_retries=5, retry_delay=2.0)

        # 属性が正しく設定されているか検証
        self.assertEqual(processor.file_paths, self.file_paths)
        self.assertEqual(processor.max_retries, 5)
        self.assertEqual(processor.retry_delay, 2.0)
        self.assertIsNone(processor.thread)
        self.assertFalse(processor.stop_event.is_set())

    @patch('threading.Thread')
    def test_start(self, mock_thread):
        """
        startメソッドの動作をテストします。

        Parameters
        ----------
        mock_thread : MagicMock
            threading.Threadのモックオブジェクト。
        """
        # スレッドのモックを設定
        mock_thread_instance = MagicMock()
        mock_thread.return_value = mock_thread_instance

        # ログが正しく出力されるか確認
        with self.assertLogs('src.excel_sync', level='INFO') as cm:
            self.processor.start()
            self.assertIn("Excel同期処理スレッドを開始しました。", cm.output[-1])

        # スレッドが正しく開始されたか検証
        mock_thread.assert_called_once_with(target=self.processor._run, daemon=True)
        mock_thread_instance.start.assert_called_once()

    @patch('threading.Thread')
    def test_stop(self, mock_thread):
        """
        stopメソッドの動作をテストします。

        Parameters
        ----------
        mock_thread : MagicMock
            threading.Threadのモックオブジェクト。
        """
        # スレッドのモックを設定
        mock_thread_instance = MagicMock()
        mock_thread_instance.is_alive.return_value = True
        self.processor.thread = mock_thread_instance

        # stopメソッドを呼び出し、適切に停止するか検証
        with patch.object(self.processor.stop_event, 'set') as mock_set, \
             self.assertLogs('src.excel_sync', level='INFO') as cm:
            self.processor.stop()
            mock_set.assert_called()
            mock_thread_instance.join.assert_called_once()
            self.assertIn("Excel同期処理スレッドを停止しました。", cm.output[-1])

    @patch('src.excel_sync.os.path.exists')
    @patch('src.excel_sync.win32com.client.DispatchEx')
    @patch('src.excel_sync.pythoncom.CoInitialize')
    @patch('src.excel_sync.pythoncom.CoUninitialize')
    @patch('src.excel_sync.time.sleep', return_value=None)
    def test_run_with_existing_files(self, mock_sleep, mock_co_uninitialize, mock_co_initialize,
                                     mock_dispatchex, mock_exists):
        """
        ファイルが存在する場合の_runメソッドの動作をテストします。

        Parameters
        ----------
        mock_sleep : MagicMock
            time.sleepのモック。
        mock_co_uninitialize : MagicMock
            pythoncom.CoUninitializeのモック。
        mock_co_initialize : MagicMock
            pythoncom.CoInitializeのモック。
        mock_dispatchex : MagicMock
            win32com.client.DispatchExのモック。
        mock_exists : MagicMock
            os.path.existsのモック。
        """
        # 全てのファイルが存在するように設定
        mock_exists.side_effect = lambda path: True

        # Excelアプリケーションとワークブックのモックを設定
        mock_excel = MagicMock()
        mock_dispatchex.return_value = mock_excel
        mock_workbook = MagicMock()
        mock_excel.Workbooks.Open.return_value = mock_workbook
        mock_excel.CalculationState = 0  # xlDone

        # _runメソッドを実行
        self.processor._run()

        # Excelアプリケーションが一度だけ起動されたか検証
        self.assertEqual(mock_dispatchex.call_count, 1)

        # 全てのファイルが処理されたか検証
        for file_path in self.file_paths:
            mock_excel.Workbooks.Open.assert_any_call(file_path)
        self.assertEqual(mock_workbook.RefreshAll.call_count, len(self.file_paths))
        self.assertEqual(mock_workbook.Save.call_count, len(self.file_paths))
        self.assertEqual(mock_workbook.Close.call_count, len(self.file_paths))

        # Excelアプリケーションが終了したか検証
        mock_excel.Quit.assert_called_once()
        mock_co_initialize.assert_called_once()
        mock_co_uninitialize.assert_called_once()

    @patch('src.excel_sync.os.path.exists')
    @patch('src.excel_sync.win32com.client.DispatchEx')
    @patch('src.excel_sync.pythoncom.CoInitialize')
    @patch('src.excel_sync.pythoncom.CoUninitialize')
    @patch('src.excel_sync.time.sleep', return_value=None)
    def test_run_with_non_existing_files(self, mock_sleep, mock_co_uninitialize, mock_co_initialize,
                                         mock_dispatchex, mock_exists):
        """
        一部のファイルが存在しない場合の_runメソッドの動作をテストします。

        Parameters
        ----------
        mock_sleep : MagicMock
            time.sleepのモック。
        mock_co_uninitialize : MagicMock
            pythoncom.CoUninitializeのモック。
        mock_co_initialize : MagicMock
            pythoncom.CoInitializeのモック。
        mock_dispatchex : MagicMock
            win32com.client.DispatchExのモック。
        mock_exists : MagicMock
            os.path.existsのモック。
        """
        # 最初のファイルが存在しないように設定
        def side_effect_exists(path):
            if settings.ACTIVITY_FILE_NAME in path:
                return False
            else:
                return True
        mock_exists.side_effect = side_effect_exists

        # Excelアプリケーションとワークブックのモックを設定
        mock_excel = MagicMock()
        mock_dispatchex.return_value = mock_excel
        mock_workbook = MagicMock()
        mock_excel.Workbooks.Open.return_value = mock_workbook
        mock_excel.CalculationState = 0  # xlDone

        # _runメソッドを実行
        self.processor._run()

        # 存在しないファイルがスキップされたか検証
        mock_exists.assert_any_call(self.file_paths[0])
        mock_exists.assert_any_call(self.file_paths[1])
        mock_exists.assert_any_call(self.file_paths[2])

        # 残りのファイルのみが処理されたか検証
        self.assertEqual(mock_excel.Workbooks.Open.call_count, 2)
        mock_excel.Workbooks.Open.assert_any_call(self.file_paths[1])
        mock_excel.Workbooks.Open.assert_any_call(self.file_paths[2])
        self.assertEqual(mock_workbook.RefreshAll.call_count, 2)
        self.assertEqual(mock_workbook.Save.call_count, 2)
        self.assertEqual(mock_workbook.Close.call_count, 2)
        mock_excel.Quit.assert_called_once()

    @patch('src.excel_sync.win32com.client.DispatchEx')
    @patch('src.excel_sync.pythoncom.CoInitialize')
    @patch('src.excel_sync.pythoncom.CoUninitialize')
    @patch('src.excel_sync.time.sleep', return_value=None)
    @patch('src.excel_sync.os.path.exists', return_value=True)
    def test_run_with_sync_failure_and_retry(self, mock_exists, mock_sleep, mock_co_uninitialize, mock_co_initialize,
                                             mock_dispatchex):
        """
        同期処理が失敗し、リトライが行われる場合の_runメソッドの動作をテストします。

        Parameters
        ----------
        mock_exists : MagicMock
            os.path.existsのモック。
        mock_sleep : MagicMock
            time.sleepのモック。
        mock_co_uninitialize : MagicMock
            pythoncom.CoUninitializeのモック。
        mock_co_initialize : MagicMock
            pythoncom.CoInitializeのモック。
        mock_dispatchex : MagicMock
            win32com.client.DispatchExのモック。
        """
        # Excelアプリケーションのインスタンスを追跡するリスト
        excel_instances = []

        # DispatchExのサイドエフェクトを設定
        def dispatchex_side_effect(*args, **kwargs):
            mock_excel = MagicMock()
            excel_instances.append(mock_excel)
            mock_workbook = MagicMock()
            mock_excel.Workbooks.Open.return_value = mock_workbook

            # RefreshAllが例外を発生させるように設定
            def refresh_all_side_effect():
                raise Exception("Test Exception during RefreshAll")

            mock_workbook.RefreshAll.side_effect = refresh_all_side_effect
            return mock_excel

        mock_dispatchex.side_effect = dispatchex_side_effect

        # _runメソッドを実行
        self.processor._run()

        # Excelアプリケーションが再起動された回数を検証
        expected_excel_instances_count = 1 + len(self.file_paths)
        self.assertEqual(len(excel_instances), expected_excel_instances_count)

        # 各ExcelインスタンスでQuitが呼ばれたか検証
        for excel_instance in excel_instances:
            excel_instance.Quit.assert_called_once()

        # DispatchExが正しい回数呼ばれたか検証
        self.assertEqual(mock_dispatchex.call_count, expected_excel_instances_count)

        # RefreshAllがリトライ回数だけ呼ばれたか検証
        total_refresh_all_calls = sum(
            [excel_instance.Workbooks.Open.return_value.RefreshAll.call_count for excel_instance in excel_instances]
        )
        expected_refresh_all_calls = len(self.file_paths) * self.processor.max_retries
        self.assertEqual(total_refresh_all_calls, expected_refresh_all_calls)

        # COMの初期化と終了が適切に行われたか検証
        mock_co_initialize.assert_called_once()
        mock_co_uninitialize.assert_called_once()

    @patch('src.excel_sync.win32com.client.DispatchEx')
    @patch('src.excel_sync.pythoncom.CoInitialize')
    @patch('src.excel_sync.pythoncom.CoUninitialize')
    @patch('src.excel_sync.time.sleep', return_value=None)
    @patch('src.excel_sync.os.path.exists', return_value=True)
    def test_run_with_stop_event(self, mock_exists, mock_sleep, mock_co_uninitialize, mock_co_initialize,
                                 mock_dispatchex):
        """
        stop_eventが設定された場合の_runメソッドの動作をテストします。

        Parameters
        ----------
        mock_exists : MagicMock
            os.path.existsのモック。
        mock_sleep : MagicMock
            time.sleepのモック。
        mock_co_uninitialize : MagicMock
            pythoncom.CoUninitializeのモック。
        mock_co_initialize : MagicMock
            pythoncom.CoInitializeのモック。
        mock_dispatchex : MagicMock
            win32com.client.DispatchExのモック。
        """
        # Excelアプリケーションとワークブックのモックを設定
        mock_workbook = MagicMock()
        mock_excel = MagicMock()
        mock_excel.Workbooks.Open.return_value = mock_workbook
        mock_excel.CalculationState = 0  # xlDone
        mock_dispatchex.return_value = mock_excel

        # ファイル処理の呼び出し回数を追跡するカウンタ
        open_call_count = [0]

        # 最初のファイル処理を通知するイベント
        processing_first_file_event = threading.Event()

        # Workbooks.Openのサイドエフェクトを設定
        def workbooks_open_side_effect(file_path):
            open_call_count[0] += 1
            if open_call_count[0] == 1:
                processing_first_file_event.set()
            return mock_workbook

        mock_excel.Workbooks.Open.side_effect = workbooks_open_side_effect

        # 同期処理を開始
        self.processor.start()

        # 最初のファイル処理が始まるのを待機
        processing_first_file_event.wait()

        # 同期処理を停止
        self.processor.stop()

        # スレッドが終了するのを待機
        self.processor.thread.join()

        # 最初のファイルのみが処理されたか検証
        self.assertEqual(open_call_count[0], 1)

        # Excelアプリケーションが終了したか検証
        mock_excel.Quit.assert_called_once()

        # stop_eventが設定されているか検証
        self.assertTrue(self.processor.stop_event.is_set())

if __name__ == '__main__':
    unittest.main()
```
## main.py
```main.py
from src.excel_sync import SynchronizedExcelProcessor
import settings
import asyncio

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

    # 他の非同期タスクを定義
    async def other_task():
        for i in range(5):
            logger.info(f"他のタスクの処理 {i+1}/5")
            await asyncio.sleep(2)  # 非同期の待機（例: I/O操作）

    # 並行して実行するタスク
    await asyncio.gather(
        other_task(),
        # 必要に応じて他の非同期タスクを追加
    )

    # Excel同期処理の完了を待つ
    while excel_processor.thread.is_alive():
        logger.info("Excel同期処理が完了するのを待っています...")
        await asyncio.sleep(1)

    logger.info("全てのタスクが完了しました。")

# 非同期イベントループの実行
if __name__ == "__main__":
    asyncio.run(main())
```