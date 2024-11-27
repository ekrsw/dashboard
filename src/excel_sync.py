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
