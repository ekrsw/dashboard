from .base_processor import BaseProcessor
import pandas as pd
import datetime
import settings
import logging

logger = logging.getLogger(__name__)

class SupportProcessor(BaseProcessor):
    def process(self):
        """
        Supportファイルのデータを指定された条件でフィルタリングおよび整形します。
        """
        try:
            base_df = self.df.fillna('') # これを入れないとdf['かんたん！保守区分'] == ''が上手く判定されない
            
            # 日付範囲でフィルタリング
            start_date = datetime.date.today()
            end_date = datetime.date.today()
            
            # start_dateからend_dateの範囲のデータを抽出
            base_df = self.filtered_by_date_range(base_df, '登録日時', start_date, end_date)

            # 直受けの検索条件
            df = base_df.copy()
            df = df[
                ((df['受付タイプ'] == '直受け') | (df['受付タイプ'] == 'HHD入電（直受け）')) & 
                (~df['顛末コード'].isin(['折返し不要・ｷｬﾝｾﾙ', 'ﾒｰﾙ・FAX回答（送信）', 'SRB投稿（要望）', 'ﾒｰﾙ・FAX文書（受信）'])) &
                ((df['かんたん！保守区分'] == '会員') | (df['かんたん！保守区分'] == '')) &
                (df['回答タイプ'] != '2次T転送')
                ]
            
            # グループ別に分割
            self.direct_ss = df[df['サポート区分'] == 'SS'].shape[0]
            self.direct_tvs = df[df['サポート区分'] == 'TVS'].shape[0]
            self.direct_kmn = df[df['サポート区分'] == '顧問先'].shape[0]
            self.direct_hhd = df[df['サポート区分'] == 'HHD'].shape[0]

            # 留守電数の検索条件
            df = base_df.copy()
            df = df[
                (df['受付タイプ'] == '留守電') & 
                (~df['顛末コード'].isin(['折返し不要・ｷｬﾝｾﾙ', 'ﾒｰﾙ・FAX回答（送信）', 'SRB投稿（要望）', 'ﾒｰﾙ・FAX文書（受信）']))
                ]
            
            # グループ別に分割
            self.ivr_ss = df[df['サポート区分'] == 'SS'].shape[0]
            self.ivr_tvs = df[df['サポート区分'] == 'TVS'].shape[0]
            self.ivr_kmn = df[df['サポート区分'] == '顧問先'].shape[0]
            self.ivr_hhd = df[df['サポート区分'] == 'HHD'].shape[0]

        except Exception as e:
            logger.error(f"直受け件数、留守電数の算定中にエラーが発生しました。: {e}")
            raise
        