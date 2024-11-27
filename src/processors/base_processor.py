import pandas as pd
import datetime
import logging
import settings

logger = logging.getLogger(__name__)

class BaseProcessor:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.df = pd.DataFrame()

    def load_data(self):
        """
        Excelファイルを読み込み、DataFrameに格納します。
        """
        try:
            self.df = pd.read_excel(self.file_path)
            logger.info(f"Loaded {self.file_path} with {self.df.shape[0]} rows.")
        except Exception as e:
            logger.error(f"Failed to load {self.file_path}: {e}")
            raise

    def save_data(self, output_file: str):
        """
        DataFrameをExcelファイルとして保存します。
        """
        try:
            self.df.to_excel(output_file, index=False)
            logger.info(f"Data saved to {output_file}.")
        except Exception as e:
            logger.error(f"Failed to save data to {output_file}: {e}")
            raise

    def process(self):
        """
        データ処理のメインメソッド。各サブクラスでオーバーライドします。
        """
        raise NotImplementedError("Subclasses should implement this method.")

    def filtered_by_date_range(self, df: pd.DataFrame, date_column: str, start_date: datetime.date, end_date: datetime.date) -> pd.DataFrame:
        """
        start_dateからend_dateの範囲のデータを抽出

        :param df: フィルタリング対象のDataFrame
        :param date_column: 登録日時のカラム名
        :param start_date: 抽出開始日
        :param end_date: 抽出終了日
        :return: フィルタリング後のDataFrame
        """
        start_date_serial = self.datetime_to_serial(datetime.datetime.combine(start_date, datetime.time.min))
        end_date_serial = self.datetime_to_serial(datetime.datetime.combine(end_date + datetime.timedelta(days=1), datetime.time.min))
        
        filtered_df = df[
            (df[date_column] >= start_date_serial) &
            (df[date_column] < end_date_serial)
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
        dt = datetime.datetime(dt.year, dt.month, dt.day)
        return (dt - base_date).total_seconds() / (24 * 60 * 60)
    
    @staticmethod
    def serial_to_datetime(serial, base_date=datetime.datetime(1899, 12, 30)) -> datetime.datetime:
        """
        シリアル値をdatetimeオブジェクトに変換する。

        :param serial: 変換するシリアル値
        :param base_date: シリアル値の基準日（デフォルトは1899年12月30日）
        :return: datetimeオブジェクト
        """
        return base_date + datetime.timedelta(days=serial)