# PyOpenRPA
This repo contains Python-based robot working on Windows which utilizes selenium and pywin32 libraries for searching and extracting debt records from FSSP service and save results in a separate file.

## NOTE! Selenium web driver was removed from this repo intentionally because the robot was published for reference purposes only.

This robot was made with Python 3.x in virtual environment deployed in Amazon AWS EC2 Windows Server instance.
_________________________________________________________________________________________________________________________________________________________________
# Welcome to fssp_extractor.py

This is a robot which purpose is to take excel file, read it row by row for debtor credentials, open fssp main page and search for debt records for every debtor, and write search results in another excel file, if any records were found.

_________________________________________________________________________________________________________________________________________________________________
# Robot structure
📦PROJECT
┣ 📦builds
┃ ┗ 📜settings.py
┃
┣ 📦docs
┃ ┣ 📜FSSP_debtor_list.xlsx
┃ ┗ 📜README.md
┃
┃ 📦resources
┃ ┣ 📂SeleniumWebDrivers
┃ ┗ 📂Chrome
┃     ┗ 📜chromedriver.exe
┃
┣ 📦sources
┃ ┣ 📂utils
┃ ┃ ┗ 📜COM_API_Excel.py
┃ ┗ 📜fssp_extractor.py
┃
┣ 📦tests
┃ ┗ 📜fssp_extractor_test.py

_________________________________________________________________________________________________________________________________________________________________

## For robot implementation, following technologies and libraries were used:

Selenium for working with Chrome instance and elements on webpage using CSS

pywin32 (win32com.client) to operate with excel as COM object

pathlib for paths composition

bas64 module to convert base64 string of captcha picture into bytes format

Image module to open, captcha picture in appropriate format

pytesseract module to recognize Russian captcha text from picture

_________________________________________________________________________________________________________________________________________________________________
## Usage:

1. Run fssp_extractor.py module directly in Python interpretator with required libraries installed


_________________________________________________________________________________________________________________________________________________________________
## What was done and what could be better:

# What was done:
1. COM_API_Excel.py separate module was created utilizing pywin32 library
2. init_webdriver() func made that correctly initializes selenium webdriver instance 
3. Necessary web elements were obtained on web page using CSS selectors
4. read_excel_row() func made that successfully reads data in excel and parses it in appropriate format 
5. search_fssp() func made that successfully searches debt records on fssp web page
6. recognize_captcha() func is working and attempts to recognize russian text on a picture, returns string
7. separate file settings.py was prepared to store all settings in one place

# What were the troubles which caused inability to accomplish task in time
1. Unable to initialize selenium webdriver instance for a long time due to restrictions in Administrator permissions on Amazon AWS EC2 Windows instance due to security reasons (error "Bluetooth: bluetooth_adapter_winrt.cc:1073 Getting Default Adapter failed")
2. Problems with text recognition on Captcha  - while recognition Unicode symbols appeared in decryption string eventually, even though it was explicitly forced to use "rus" library - probably because Russian recognition language package is not very enriched due to limited usage of pytesseract with Russian language

_________________________________________________________________________________________________________________________________________________________________
## Documentation used:

pywin32
---------------
https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb178833(v=office.12)?redirectedfrom=MSDN 
http://timgolden.me.uk/python/win32_how_do_i/generate-a-static-com-proxy.html 
https://stackoverflow.com/questions/39877278/python-open-excel-workbook-using-win32-com-api 
http://snakeproject-ru.1gb.ru/rubric/article.php?art=python_win32com_client 
http://raaviblog.com/python-2-7-read-and-write-excel-file-with-win32com/ 


pytesseract
---------------
https://github.com/tesseract-ocr/tesseract 
https://github.com/tesseract-ocr/tessdata
https://tesseract-ocr.github.io/tessdoc/Home.html#binaries
https://stackoverflow.com/questions/7391945/how-do-i-read-image-data-from-a-url-in-python 


pathlib
---------------
https://docs.python.org/3/library/pathlib.html#pathlib.Path.is_file


Image module
---------------
https://pillow.readthedocs.io/en/stable/reference/Image.html

