from pathlib import Path

cwd = Path(__file__).absolute()

global_settings = {
    "chrome": {
        "chrome_portable_exe_path": cwd / Path(r'../resources/GoogleChromePortable/GoogleChromePortable.exe').absolute(),
        "chrome_exe_path": Path(r'C:\Program Files\Google\Chrome\Application\chrome.exe'),
        "chrome_driver_path": cwd / Path(r'../resources/SeleniumWebDrivers/Chrome/chromedriver.exe').absolute()},
    "excel": {
        "excel_file_path": cwd / Path(r'../docs/FSSP_debtor_list.xlsx').absolute(),
        "excel_configuration": {"worksheet_index_int": 1,
                                "from_row_int": 2,
                                "from_col_int": 1
                                }
    },
    "pytesseract_binary_file_path": r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    "fssp": {
        "fssp_web_url": "https://fssprus.ru/",
        "fssp_popup_window_css_selector": 'div[class="modal-info"]',
        "fssp_popup_window_close_css_selector": 'button[class="tingle-modal__close"]',
        "fssp_input_box_css_selector": 'input[id="debt-form01"]',
        "fssp_search_button_css_selector": 'button[class="btn btn-primary"]',
        "fssp_captcha_pic_css_selector": 'img[id="capchaVisual"]',
        "fssp_captcha_input_box_css_selector": 'input[id="captcha-popup-code"]',
        "fssp_captcha_send_button_css_selector": 'input[class="input-submit-capcha"]',
        "fssp_captcha_error_css_selector": 'div[class="b-form__label b-form__label--error"]'

    },
    "sudrf": {
        "sudrf_web_url": "https://sudrf.ru/index.php?id=300#sp"
    }
}
