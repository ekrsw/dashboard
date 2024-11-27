import pandas as pd
from typing import Any, List
import logging

def read_excel_data(file_path: str, sheet_name: str = 0, columns_to_keep: List[str] = None, date_columns: List[str] = None) -> pd.DataFrame:
    """
    Excelファイルからデータを読み込み、不要なカラムを削除し、日時をDatetime型に変換する関数。

    :param file_path: Excelファイルのパス
    :param sheet_name: シート名またはインデックス
    :param columns_to_keep: 残したいカラムのリスト
    :param date_columns: Datetime型に変換するカラムのリスト
    :return: pandas DataFrame
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        logging.info(f"Excelデータを読み込みました: {file_path}")
        
        if columns_to_keep:
            missing_cols = [col for col in columns_to_keep if col not in df.columns]
            if missing_cols:
                logging.error(f"指定されたカラムが存在しません: {missing_cols}")
                raise KeyError(f"Missing columns: {missing_cols}")
            df = df[columns_to_keep]
            logging.info(f"不要なカラムを削除しました。残されたカラム: {columns_to_keep}")
        
        if date_columns:
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                    logging.info(f"指定されたカラムをDatetime型に変換しました: {col}")
                else:
                    logging.warning(f"指定された日時カラムが存在しません: {col}")
        
        return df
    except Exception as e:
        logging.error(f"Excelデータの読み込みに失敗しました: {e}")
        raise
