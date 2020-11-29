
from inspect import trace
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import requests
import configparser
import log
import openpyxl
import traceback
import os




# logフォルダなければ作成
config = configparser.ConfigParser()
config.read(R"setting.ini",encoding="utf-8")

log_path = os.path.abspath('src/log/debug.log')
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


def create_search_querystring(search_colors,is_grayscale):
    """画像検索用のQueryString生成

    Args:
        search_colors ([type]): 色
        is_grayscale (bool): グレースケール

    Returns:
        [type]: [description]
    """

    logger.debug("-------検索用QueryString生成処理開始")

    search_colors = search_colors.split(",")

    query_string = "?"
    for search_color in search_colors:

        query_string += "colors=" + search_color +"&"
        
    query_string = query_string[:-1]
    logger.debug("生成されたQueryString : " + query_string)

    logger.debug("-------検索用QueryString生成完了")

    return query_string



def save_image(browser,output_path,search_word):
    """画像保存

    Args:
        browser ([type]): [description]
        output_path ([type]): [description]
        search_word ([type]): [description]
    """

    logger.debug("-------画像保存処理開始")
    image_dom = browser.find_element_by_id("media_container")
    save_folder_path = output_path + search_word

    if not os.path.exists(save_folder_path):
        os.makedirs(save_folder_path)

    img_url = image_dom.find_element_by_tag_name.get_attribute("img").get_attribute("src")
    file_name = img_url.split("/")[-1]

    logger.debug("画像URL : " + img_url)
    logger.debug("ファイル名 : " + file_name)

    request = requests.get(img_url)
    with open(save_folder_path + '\\' + file_name,'wb') as f:
        f.write(request.content)


    logger.debug("-------画像保存処理完了")

    

if __name__ == "__main__":


    # webdriver起動 
    chromedriver_path = os.path.abspath(config.get("path","chromedrvier") )
    browser = webdriver.Chrome(executable_path=chromedriver_path)   

    pixabay_base_url = config.get("url","pixabay_search_url") # 検索ベースURL
    input_excel_path = os.path.abspath(config.get("path","input_excel")) # inputExcelパス
    search_word_cell = config.get("excel","search_word_cell") # 検索ワードセル
    search_colors = config.get("searchParameter","colors")
    is_grayscale = config.get("searchParameter","is_grayscale")
    output_folder_path = os.path.abspath(config.get("path","outpu_folder"))


    logger.debug("---------------------------------------------")
    logger.debug("chromedriverパス : " + chromedriver_path)
    logger.debug("pixabay検索URL : " + pixabay_base_url)
    logger.debug("画像保存フォルダーパス : " + output_folder_path)
    logger.debug("インプットexcelパス : " + input_excel_path)
    logger.debug("検索ワードセル : " + search_word_cell)
    logger.debug("検索色 : " + search_colors)
    logger.debug("グレースケール : " + is_grayscale)
    logger.debug("---------------------------------------------")


    try:
        search_word = get_search_word(input_excel_path,search_word_cell)

        logger.debug("検索ワード : " + search_word)
        logger.debug("検索URL : " + pixabay_base_url + search_word)

        
        query_string = create_search_querystring(search_colors,is_grayscale)

        # 
        browser.get(pixabay_base_url + search_word + query_string)
        wait_browser(browser)

        logger.debug("pixabayに接続完了")

        
        # 画像保存のために新しいタブで各画像ページを開く
        search_result_dom = browser.find_element_by_class_name("search_results")
        images_dom = search_result_dom.find_elements_by_class_name("item")
        for image_dom in images_dom:
            page_url = image_dom.get_attribute("href")
            
            # 新しいwindow
            browser.execute_script("window.open()")
            browser.switch_to.window(browser.window_handles[-1])
            browser.get(page_url)
            wait_browser(browser)

            # 画像保存
            save_image(browser,output_folder_path,search_word)

            browser.close()
            browser.switch_to.window(browser.window_handles[0])
    except Exception:
        logger.error(traceback.format_exc())

    