import win32com.client as win32
import datetime  # Datetime functions


def init_excel():
    """
    Init excel com instance

    :return excel: com object of excel  
    """
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    excel.DisplayAlerts = False
    return excel


def destruct_excel(excel):
    """
    Close excel com instance

    :param excel: com object of excel  
    """
    excel.DisplayAlerts = True
    excel.Quit()
    excel = None


def open_workbook(excel, file_path: str):
    """
    Opens Excel workbook from specified path

    :param excel: com object of excel  
    :param str file_path: 
    """
    excel.AskToUpdateLinks = False
    try:
        xlwb = excel.Workbooks(file_path)
    except Exception as e:
        try:
            xlwb = excel.Workbooks.Open(file_path)
        except Exception as e:
            excel.AskToUpdateLinks = True
            raise e
    excel.AskToUpdateLinks = True
    return xlwb


def get_count_worksheet(workbook) -> int:
    """
    Get worksheet count. Index from 1

    :param workbook: com object of excel  
    """
    return workbook.Worksheets.Count


def save_worksheet(workbook, file_path: str):
    """
    Save excel workbook

    :param workbook: com object of excel  
    :param str file_path: path to saving file
    """
    workbook.SaveAs(Filename=file_path)


def close_workbook(workbook, savebool=True):
    """
    Save the workbook

    :param workbook: com object of excel  
    :param bool savebool: default True
    """
    workbook.Close(savebool)


def read_range(workbook, worksheet_index_int: int, from_row_int: int, from_col_int: int,
               to_row_int: int, to_col_int: int) -> tuple:
    """
    Read range from XLS

    :param workbook: com object of excel        
    :param int worksheet_index_int: index of sheet (start by 1)
    :param int from_row_int: 
    :param int from_col_int: 
    :param int to_row_int: 
    :param int to_col_int: 
    :return tuple result: turple of turples of the cell value - None if cell is empty
    """
    result = None
    sheet = workbook.Worksheets(worksheet_index_int)
    result = sheet.Range(sheet.Cells(from_row_int, from_col_int), sheet.Cells(to_row_int, to_col_int)).Value
    return result


def set_range(workbook, worksheet_index_int: int, from_row_int, from_col_int, data_list: list):
    """
    Set value in range

    :param workbook: com object of excel 
    :param int worksheet_index_int: index of sheet (start by 1)
    :param int from_row_int:
    :param int from_col_int: 
    :param list data_list: 
    """
    sheet = workbook.Worksheets(worksheet_index_int)
    sheet.Range(sheet.Cells(from_row_int, from_col_int), sheet.Cells(from_row_int, from_col_int)).Value = data_list


def convert_excel_datetime_to_python_datetime(excel_datetime):
    """
    Convert excel datetime to python datetime

    :param excel_datetime: 
    :return python_datetime:
    """
    python_datetime = None
    try:
        python_datetime = datetime.datetime(year=excel_datetime.year, month=excel_datetime.month,
                                            day=excel_datetime.day, hour=excel_datetime.hour,
                                            minute=excel_datetime.minute, second=excel_datetime.second)
    except Exception as e:
        pass
    return python_datetime


def convert_excel_date_to_python_date(excel_date):
    """
    Convert excel date to python date

    :param excel_date:
    :return python_date: 
    """
    python_date = None
    try:
        python_date = datetime.date(year=excel_date.year, month=excel_date.month, day=excel_date.day)
    except Exception as e:
        pass
    return python_date


def count_rows(workbook, worksheet_index_int: int):
    """
    Count number of non-empty rows in a column

    :param workbook: com object excel 
    :param int worksheet_index_int: index of sheet (start by 1)
    :return int: 
    """
    sheet = workbook.Worksheets(worksheet_index_int)
    used = sheet.UsedRange
    nrows = used.Row + used.Rows.Count - 1
    return nrows


def count_columns(workbook, worksheet_index_int: int):
    """
    Count number of non-empty columns in a worksheet

    :param workbook: com object excel 
    :param int worksheet_index_int: index of sheet (start by 1)
    :return int: 
    """
    sheet = workbook.Worksheets(worksheet_index_int)
    used = sheet.UsedRange
    ncols = used.Column + used.Columns.Count - 1
    return ncols


def insert_to_cell(workbook, worksheet_index_int: int, row_num: int, col_num: int, paste_text: str):
    """
    Insert data in a particular cell in a worksheet

    :param workbook: com object excel 
    :param int worksheet_index_int: index of sheet (start by 1)
    """
    sheet = workbook.Worksheets(worksheet_index_int)
    sheet.Cells(row_num, col_num).Value = paste_text
