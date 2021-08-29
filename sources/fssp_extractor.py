import time
import logging
import base64
import keyboard
import pytesseract

from PIL import Image
from io import BytesIO
from pathlib import Path
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException

from builds.settings import global_settings
from .utils import COM_API_Excel

#  setup separate logger for current module
logger = logging.getLogger(__name__)

#  initialize explicit path to Tesseract binary file
pytesseract.pytesseract.tesseract_cmd = global_settings["pytesseract_binary_file_path"]


def init_webdriver(chrome_exe_path, chrome_driver_path=None) -> webdriver:
    """
    Initialize selenium webdriver instance with required options
    :param Path chrome_exe_path: absolute path to portable Chrome executable
    :param Path chrome_driver_path: absolute path to Chrome driver executable
    :return: selenium webdriver instance
    """
    # establish options instance and apply settings
    options_instance = webdriver.ChromeOptions()
    options_instance.binary_location = str(chrome_exe_path)

    options_instance.add_argument("--start-maximized")
    options_instance.add_argument("--remote-debugging-port=9222")
    options_instance.add_argument("--ignore-certificate-errors")
    options_instance.add_argument("--disable-extensions")
    options_instance.add_argument("--no-sandbox")
    options_instance.add_argument("--disable-dev-shm-usage")
    options_instance.add_experimental_option("excludeSwitches", ["enable-logging"])
    driver_instance = None  # initialize driver instance
    if chrome_driver_path:
        driver_instance = webdriver.Chrome(executable_path=str(chrome_driver_path), options=options_instance)
    else:
        driver_instance = webdriver.Chrome(options=options_instance)
    return driver_instance


def read_excel_row(excel_file_path: Path, worksheet_index_int: int, from_row_int: int, from_col_int: int, logger=None) -> list:
    """
    Opens Excel file, reads debtor info by row
    :param Path excel_file_path: file path to Excel book with debtor data
    :param int worksheet_index_int: index of excel worksheet (starts from 1) 
    :param int from_row_int: index of row in excel (starts from 1)
    :param int from_col_int: index of column in excel (starts from 1)
    :return: list
    """
    excel = COM_API_Excel.init_excel()  # initialize excel instance
    try:
        workbook = COM_API_Excel.open_workbook(excel, str(excel_file_path))  # open excel file
        ncols = COM_API_Excel.count_columns(workbook, worksheet_index_int)  # count num of columns
        row_data = COM_API_Excel.read_range(workbook, worksheet_index_int,  # read row and return as tuple
                                            from_row_int=from_row_int, from_col_int=from_col_int,
                                            to_row_int=from_row_int, to_col_int=ncols)
        surname, name, patronym, birth_date, *args = [element for row in row_data for element in row]  #
        birth_date = COM_API_Excel.convert_excel_datetime_to_python_datetime(birth_date).strftime("%d.%m.%Y")
        debtor_data = [surname, name, patronym, birth_date]
        return debtor_data
    except Exception as e:
        if logger: logger.warning(f'Unable to extract debtor data from Excel file:{e}')
    finally:
        COM_API_Excel.destruct_excel(excel)


def search_fssp(driver_instance, debtor_data: tuple, fssp_web_url: str, fssp_popup_window_css_selector: str,
                fssp_popup_window_close_css_selector: str, fssp_input_box_css_selector: str,
                fssp_search_button_css_selector: str, logger=None) -> None:
    """
    Opens fssp main page and searches records by debtor data
    :param driver_instance: selenium webdriver instance
    :param list debtor_data: debtor credentials required to search
    :param str fssp_web_url: url address of fssp main page
    :param str fssp_popup_window_css_selector: 
    :param str fssp_popup_window_close_css_selector: 
    :param str fssp_input_box_css_selector: 
    :param str fssp_search_button_css_selector: 
    :return: None
    """
    driver_instance.get(fssp_web_url)  # open fssp web page
    driver_instance.implicitly_wait(5)
    try:
        popup_window = driver_instance.find_element_by_css_selector(fssp_popup_window_css_selector)
        popup_window_close = driver_instance.find_element_by_css_selector(fssp_popup_window_close_css_selector)
        if popup_window:  # check if popup window appeared in browser
            popup_window_close.click()  # close popup window
    except NoSuchElementException as e:
        if logger: logger.info(e)
    # input debtor data into search box and click "search"
    input_box = driver_instance.find_element_by_css_selector(fssp_input_box_css_selector)
    input_box.send_keys(" ".join(debtor_data))  # concatenate list of str to search string
    time.sleep(0.3)
    search_button = driver_instance.find_element_by_css_selector(fssp_search_button_css_selector)
    search_button.click()
    driver_instance.implicitly_wait(10)  # waiting either until page loads fully or 10 sec whichever sooner
    return None


def recognize_captcha(driver_instance, fssp_captcha_pic_css_selector: str, fssp_captcha_input_box_css_selector: str,
                      fssp_captcha_send_button_css_selector: str, fssp_captcha_error_css_selector: str, logger=None) -> None:
    """
    Attempts to recognize captcha and paste it into input box
    :param driver_instance: selenium webdriver instance
    :param str fssp_captcha_pic_css_selector: 
    :param str fssp_captcha_input_box_css_selector: 
    :param str fssp_captcha_send_button_css_selector: 
    :param str fssp_captcha_error_css_selector: 
    :param logger: logger
    :return: None
    """
    try:
        captcha_input_box = driver_instance.find_element_by_css_selector(fssp_captcha_input_box_css_selector)
        captcha_send_button = driver_instance.find_element_by_css_selector(fssp_captcha_send_button_css_selector)
        captcha_pic = driver_instance.find_element_by_css_selector(fssp_captcha_pic_css_selector)
        captcha_pic_src = captcha_pic.get_attribute('src')  # extract src attribute which is base64 string of captcha pic
        captcha_pic_base64 = captcha_pic_src[23:]  # crop first 23 symbols to get base64 string of captcha image
        decrypted_string = ""
        try:
            captcha_pic_decode = base64.b64decode(captcha_pic_base64)  # decode base64 to bytes representation
            captcha_pic_jpg = Image.open(BytesIO(captcha_pic_decode))  # initialize Image object from captcha pic
            decrypted_string = pytesseract.image_to_string(captcha_pic_jpg, lang="rus")
        except Exception as e:
            if logger: logger.exception(f"Error while captcha decryption: {e}")

        captcha_input_box.send_keys(decrypted_string)
        captcha_send_button.click()
        try:
            captcha_error = driver_instance.find_element_by_css_selector(fssp_captcha_error_css_selector)
            if captcha_error:  # if captcha was not entered correctly
                    keyboard.send("ctrl+r")  # reload page
                    driver_instance.implicitly_wait(10)
        except NoSuchElementException as e:
            if logger: logger.info(e)
    except Exception as e:
        if logger: logger.warning(f'Error when trying to extract captcha data:{e}')
    return None


def extract_data_from_fssp(driver_instance, logger=None) -> list:
    """
    Extracts data from FSSP results page and returns as list
    :param driver_instance: selenium webdriver instance
    :param logger: logger
    :return: list
    """
    return None


def paste_fssp_data_excel(excel_file_path: Path, debtor_data: list, worksheet_index_int: int, from_row_int: int, from_col_int: int,
                          logger=None) -> None:
    """
    Pastes data to new Excel workbook if any results were found for particular row and saves excel file
    :param Path excel_file_path: file path to Excel book with debtor data
    :param list debtor_data: data extracted for particular debtor from fssp web page 
    :param int worksheet_index_int: index of excel worksheet (starts from 1) 
    :param int from_row_int: index of row in excel (starts from 1)
    :param int from_col_int: index of column in excel (starts from 1)
    :param logger: logger
    :return: None
    """
    return None


# driver code
if __name__ == "__main__":
    # initialize webdriver instance
    driver_instance = init_webdriver(
        chrome_exe_path=global_settings["chrome"]["chrome_exe_path"],
        chrome_driver_path=global_settings["chrome"]["chrome_driver_path"])
    # extract debtor credentials from excel workbook
    debtor_data = read_excel_row(
        excel_file_path=global_settings["excel"]["excel_file_path"],
        worksheet_index_int=global_settings["excel"]["excel_configuration"]["worksheet_index_int"],
        from_row_int=global_settings["excel"]["excel_configuration"]["from_row_int"],
        from_col_int=global_settings["excel"]["excel_configuration"]["from_col_int"], logger=logger)
    # paste data and search for results in fssp web interface
    search_fssp(driver_instance, debtor_data,
                               fssp_web_url=global_settings["fssp"]["fssp_web_url"],
                               fssp_popup_window_css_selector=global_settings["fssp"]["fssp_popup_window_css_selector"],
                               fssp_popup_window_close_css_selector=global_settings["fssp"][
                                   "fssp_popup_window_close_css_selector"],
                               fssp_input_box_css_selector=global_settings["fssp"]["fssp_input_box_css_selector"],
                               fssp_search_button_css_selector=global_settings["fssp"][
                                   "fssp_search_button_css_selector"])
    # attempt to recognize captcha
    recognize_captcha(driver_instance, fssp_captcha_pic_css_selector=global_settings["fssp_captcha_pic_css_selector"],
                      fssp_captcha_input_box_css_selector=global_settings["fssp_captcha_input_box_css_selector"],
                      fssp_captcha_send_button_css_selector=global_settings["fssp_captcha_send_button_css_selector"],
                      fssp_captcha_error_css_selector=global_settings["fssp_captcha_error_css_selector"], logger=logger)
    # extract debtor data from search results page
    debtor_data = extract_data_from_fssp(driver_instance, logger)
    # paste data, if any, to another excel workbook
    paste_fssp_data_excel(excel_file_path, debtor_data, worksheet_index_int, from_row_int, from_col_int, logger)
