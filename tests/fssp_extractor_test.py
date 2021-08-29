import logging
import keyboard
from sources import fssp_extractor
from builds.settings import global_settings

logger = logging.getLogger(__name__)


if __name__ == "__main__":
    result = []
    driver_instance = fssp_extractor.init_webdriver(
        chrome_exe_path=global_settings["chrome"]["chrome_exe_path"],
        chrome_driver_path=global_settings["chrome"]["chrome_driver_path"])
    debtor_data = fssp_extractor.read_excel_row(
        excel_file_path=global_settings["excel"]["excel_file_path"],
        worksheet_index_int=global_settings["excel"]["excel_configuration"]["worksheet_index_int"],
        from_row_int=global_settings["excel"]["excel_configuration"]["from_row_int"],
        from_col_int=global_settings["excel"]["excel_configuration"]["from_col_int"], logger=logger)
    fssp_extractor.search_fssp(driver_instance, debtor_data,
                               fssp_web_url=global_settings["fssp"]["fssp_web_url"],
                               fssp_popup_window_css_selector=global_settings["fssp"]["fssp_popup_window_css_selector"],
                               fssp_popup_window_close_css_selector=global_settings["fssp"][
                                   "fssp_popup_window_close_css_selector"],
                               fssp_input_box_css_selector=global_settings["fssp"]["fssp_input_box_css_selector"],
                               fssp_search_button_css_selector=global_settings["fssp"][
                                   "fssp_search_button_css_selector"])

    # try_counter = 10
    # while try_counter:
    #     fssp_extractor.recognize_captcha(driver_instance, logger=logger,
    #                                      fssp_captcha_pic_css_selector=global_settings["fssp_captcha_pic_css_selector"],
    #                                      fssp_captcha_input_box_css_selector=global_settings[
    #                                          "fssp_captcha_input_box_css_selector"],
    #                                      fssp_captcha_send_button_css_selector=global_settings[
    #                                          "fssp_captcha_send_button_css_selector"],
    #                                      fssp_captcha_error_css_selector=global_settings[
    #                                          "fssp_captcha_error_css_selector"])
    #     try_counter -= 1

    result.append(debtor_data)
    print(result)
