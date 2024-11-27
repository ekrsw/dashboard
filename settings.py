import os
from dotenv import load_dotenv

# .envファイルの読み込み
load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Excelファイルの名前とパス
ACTIVITY_FILE_NAME = 'TS_todays_activity.xlsx'
CLOSE_FILE_NAME = 'TS_todays_close.xlsx'
SUPPORT_FILE_NAME = 'TS_todays_support.xlsx'


ACTIVITY_FILE = os.path.join(BASE_DIR, 'data', ACTIVITY_FILE_NAME)
CLOSE_FILE = os.path.join(BASE_DIR, 'data', CLOSE_FILE_NAME)
SUPPORT_FILE = os.path.join(BASE_DIR, 'data', SUPPORT_FILE_NAME)

EXCEL_FILES = [ACTIVITY_FILE, CLOSE_FILE, SUPPORT_FILE]

# Excel同期処理の設定
SYNC_MAX_RETRIES = 5  # 同期失敗時の最大リトライ回数
SYNC_RETRY_DELAY = 2.0  # リトライ間の待機時間（秒）
REFRESH_INTERVAL = 5  # 更新が完了するまで待機する時間（秒）

# 同期処理で保持するカラム
COLUMNS_TO_KEEP = {
    'TS_todays_activity.xlsx': ['Revenue', 'Cost', 'Date'],
    'TS_todays_close.xlsx': ['Revenue', 'Cost', 'Date'],
    'TS_todays_support.xlsx': ['Revenue', 'Cost', 'Date']
}

# Datetime型に変換するカラム
DATE_COLUMNS = ['Date']

# Webスクレイピングの設定
WEB_URL = 'https://example.com/data'  # 実際のURLに置き換えてください

# KPIの出力先
OUTPUT_CSV = 'data/output/kpis.csv'

# ログファイルのパス
LOG_FILE = 'kpi_dashboard.log'

# テストコード
TEST_DATA_PATH = os.path.join(os.path.join(BASE_DIR, 'tests'), 'test_data')

# シリアル値
SERIAL_20_MINUTES = 0.0138888888888889
SERIAL_30_MINUTES = 0.0208333333333333
SERIAL_40_MINUTES = 0.0277777777777778
SERIAL_60_MINUTES = 0.0416666666666667

# CTStageレポーター関係設定
REPORTER_URL = os.getenv('REPORTER_URL')
REPORTER_ID = os.getenv('REPORTER_ID')
HEADLESS_MODE = False
RETRY_COUNT = 3
DELAY = 3
TEMPLATE_SS = ['パブリック', '対応状況集計表用-SS']
TEMPLATE_TVS = ['パブリック', '対応状況集計表用-TVS']
TEMPLATE_KMN = ['パブリック', '対応状況集計表用-顧問先']
TEMPLATE_HHD = ['パブリック', '対応状況集計表用-HHD']