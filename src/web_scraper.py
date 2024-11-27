import requests
from bs4 import BeautifulSoup
from typing import Any
import logging

def fetch_web_data(url: str, params: dict = None) -> Any:
    """
    Webページからデータを取得する関数。

    :param url: データを取得するURL
    :param params: 必要に応じたクエリパラメータ
    :return: 取得したデータ（BeautifulSoupオブジェクトなど）
    """
    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        logging.info(f"Webデータを取得しました: {url}")
        return soup
    except requests.RequestException as e:
        logging.error(f"Webデータの取得に失敗しました: {e}")
        raise
