
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from time import sleep

import configparser
from src import log
import  openpyxl

import os




# logフォルダなければ作成
config = configparser.ConfigParser()
config.read(R"setting.ini",encoding="utf-8")

log_path = '../log/debug.log'
if not os.path.exists(log_path):
    os.mkdir(log_path)
logger = log.Logger(__name__)






def wait_browser(browser,wait_time=0):
    """ブラウザ待機

    Args:
        browser ([type]): [description]
        wait_time (int, optional): [description]. Defaults to 0.
    """
    # ページ読み込みまで待機 かつ１秒待機
    # 15秒待機でタイムアウト
    WebDriverWait(browser, 15).until(EC.presence_of_all_elements_located)
    sleep(wait_time)


def get_search_word(input_excel_path,search_word_cell):
    """Excelから検索ワードを取得する

    Args:
        input_excel_path ([type]): [description]

    Returns:
        [type]: [description]
    """
    logger.debug("------Excelから検索ワード取得処理開始")

    search_word = ""

    wb = openpyxl.load_workbook(input_excel_path)
    ws = wb.worksheets[0]

    search_word = ws.cell[search_word_cell]


    logger.debug("------Excelから検索ワード取得処理開始")

    return search_word


def create_search_querystring():

    logger.debug("-------検索用QueryString生成処理開始")


    search_colors = config.get("searchParameter","colors")
    is_grayscale = config.get("searchParameter","is_grayscale")


    logger.debug("検索色 : " + search_colors)
    logger.debug("グレースケール : " + is_grayscale)

    search_colors = search_colors.split(",")
    


    
    logger.debug("-------検索用QueryString生成完了")

if __name__ == "__main__":


    # webdriver起動 
    chromedriver_path = os.path.abspath(config.get("path","chromedrvier") )
    browser = webdriver.Chrome(executable_path=chromedriver_path)   

    pixabay_base_url = config.get("url","pixabay_search_url") # 検索ベースURL
    input_excel_path = os.path.abspath(config.get("path","input_excel")) # inputExcelパス
    search_word_cell = config.get("excel","search_word_cell") # 検索ワードセル


    logger.debug("---------------------------------------------")
    logger.debug("chromedriverパス : " + chromedriver_path)
    logger.debug("pixabay検索URL : " + pixabay_base_url)
    logger.debug("インプットexcelパス : " + input_excel_path)
    logger.debug("検索ワードセル : " + search_word_cell)
    logger.debug("---------------------------------------------")


    search_word = get_search_word(input_excel_path,search_word_cell)

    logger.debug("検索ワード : " + search_word)
    logger.debug("検索URL : " + pixabay_base_url + search_word)

    browser.get(pixabay_base_url + search_word)
    wait_browser(browser)

    logger.debug("pixabayに接続完了")
