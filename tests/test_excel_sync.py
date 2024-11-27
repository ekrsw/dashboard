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
        self.processor = SynchronizedExcelProcessor(self.file_paths, max_retries=0, retry_delay=0.1)

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
