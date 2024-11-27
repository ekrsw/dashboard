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
            self.df = self.filtered_by_date_range(self.df, '登録日時 (関連) (サポート案件)', start_date, end_date)
            logger.info(f"Filtered DataFrame by date range {start_date} to {end_date}. Remaining rows: {self.df.shape[0]}")

            # 案件番号でソートし、最も早い日時を残して重複を削除
            self.df = self.df.sort_values(by=['案件番号 (関連) (サポート案件)', '登録日時'])
            self.df = self.df.drop_duplicates(subset='案件番号 (関連) (サポート案件)', keep='first')
            logger.info(f"Sorted and dropped duplicates. Remaining rows: {self.df.shape[0]}")

            # '登録日時'と'登録日時 (関連) (サポート案件)'をdatetime型に変換
            # self.df['登録日時'] = pd.to_datetime(self.df['登録日時'], unit='d', origin='1899-12-30')
            # self.df['登録日時 (関連) (サポート案件)'] = pd.to_datetime(self.df['登録日時 (関連) (サポート案件)'], unit='d', origin='1899-12-30')

            # 時間差を計算
            self.df['時間差'] = self.df['登録日時'] - self.df['登録日時 (関連) (サポート案件)']
            self.df['時間差'] = self.df['時間差'].fillna(0.0)
            logger.info("Calculated '時間差' column.")

            # 受付タイプが折返しor留守電のものを抽出
            _df_ss_tvs_kmn = self.df[(self.df['受付タイプ (関連) (サポート案件)'] == '折返し') | (self.df['受付タイプ (関連) (サポート案件)'] == '留守電')]
            _df_hhd = self.df[(self.df['受付タイプ (関連) (サポート案件)'] == 'HHD入電（折返し）') | (self.df['受付タイプ (関連) (サポート案件)'] == '留守電')]

            df_ss = _df_ss_tvs_kmn[(_df_ss_tvs_kmn['サポート区分 (関連) (サポート案件)'] == 'SS')]
            df_tvs = _df_ss_tvs_kmn[(_df_ss_tvs_kmn['サポート区分 (関連) (サポート案件)'] == 'TVS')]
            df_tvs.to_excel('df_tvs.xlsx')
            df_kmn = _df_ss_tvs_kmn[(_df_ss_tvs_kmn['サポート区分 (関連) (サポート案件)'] == '顧問先')]
            df_hhd = _df_hhd[(_df_hhd['サポート区分 (関連) (サポート案件)'] == 'HHD')]

            # グループ別に分割
            self.cb_0_20_ss, self.cb_20_30_ss, self.cb_30_40_ss, self.cb_40_60_ss, self.cb_60over_ss, self.cb_not_include_ss = self.group_activities_by_callback_duration(df_ss)
            self.cb_0_20_tvs, self.cb_20_30_tvs, self.cb_30_40_tvs, self.cb_40_60_tvs, self.cb_60over_tvs, self.cb_not_include_tvs = self.group_activities_by_callback_duration(df_tvs)
            self.cb_0_20_kmn, self.cb_20_30_kmn, self.cb_30_40_kmn, self.cb_40_60_kmn, self.cb_60over_kmn, self.cb_not_include_kmn = self.group_activities_by_callback_duration(df_kmn)
            self.cb_0_20_hhd, self.cb_20_30_hhd, self.cb_30_40_hhd, self.cb_40_60_hhd, self.cb_60over_hhd, self.cb_not_include_hhd = self.group_activities_by_callback_duration(df_hhd)

        except Exception as e:
            logger.error(f"Error processing activity data: {e}")
            raise
    
    def group_activities_by_callback_duration(self, df):
        towenty_minutes = settings.SERIAL_20_MINUTES
        thirty_minutes = settings.SERIAL_30_MINUTES
        forty_minutes = settings.SERIAL_40_MINUTES
        sixty_minutes = settings.SERIAL_60_MINUTES

        cb_0_20 = df[(df['時間差'] <= towenty_minutes)].shape[0]
        cb_20_30 = df[(df['時間差'] > towenty_minutes) & (df['時間差'] <= thirty_minutes)].shape[0]
        cb_30_40 = df[(df['時間差'] > thirty_minutes) & (df['時間差'] <= forty_minutes)].shape[0]
        cb_40_60 = df[(df['時間差'] > forty_minutes) & (df['時間差'] <= sixty_minutes) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')].shape[0]
        cb_60over = df[(df['時間差'] > sixty_minutes) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')].shape[0]
        cb_not_include =df[(df['時間差'] > sixty_minutes) & (df['指標に含めない (関連) (サポート案件)'] == 'はい')].shape[0]

        return cb_0_20, cb_20_30, cb_30_40, cb_40_60, cb_60over, cb_not_include

    @staticmethod
    def datetime_to_serial(dt: datetime.datetime, base_date=datetime.datetime(1899, 12, 30)) -> float:
        """
        datetimeオブジェクトをシリアル値に変換する。

        :param dt: 変換するdatetimeオブジェクト
        :param base_date: シリアル値の基準日（デフォルトは1899年12月30日）
        :return: シリアル値
        """
        return (dt - base_date).total_seconds() / (24 * 60 * 60)
