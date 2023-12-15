r"""
Generate raw report for each student.

Example
-------
::
    >>> import raw

        Split the monthly report for each student.
            data_path: sleected excel file.
            raw_path: monthly_report/[...]/raw_reports.

    >>> raw.split_book(data_path, raw_path)

        Render the format for each individual report.
    
    >>> raw.format_render(raw_path)

Author: Rayykia

Date: 12-6-2023
"""

import pandas as pd
import xlwings as xw
import glob


import utils

from tqdm import tqdm


def split_book(
        data: pd.DataFrame,
        o_dir: str
) -> None:
    """Split the monthly report for each student.
    
    Example
    -------
    >>> split_book(data, r'./raw_report') 

    Parameters
    ----------
    i_path: str or path-like object
        path of the report
    o_dir: str or path-like object
        the output directory

    .. note:
        Might need to update if the format of the monthly report changes as the 
        raw report for each student are picked the first 8 columns of the montly
        report.
    """

    
    student_list = data.iloc[:, 0].unique()


    print("Generating raw reports...")
    for student in tqdm(student_list):
        output_path = "/".join([o_dir, '{}.xlsx'.format(student)])
        data[data.iloc[:, 0].isin([student])].to_excel(output_path, index=False)
    print("Raw reports generated.\n")




def _newline_render(
        text: str
) -> str:
    """Replace multiple newlines with a sigle newline.
    
    Example
    -------
    >>> rendered_text = _newline_render(text)

    Parameters
    ----------
    text: str
        the text needs to be rendered
    """
    text = text.replace("\n\n", "\n")
    text = text.replace("\n\n", "\n")
    return text


def _single_text_render(
        text: str,
        max_characters: int,
) -> str:
    len_text= len(text)
    n = len_text // 20
    for i in range(1, n+1):
        if len_text > max_characters*i:
            split_index = max_characters*i  # Find the last space within the character limit
            if split_index != len_text:
                text = text[:split_index] + '\n' + text[split_index+1:]
    return text

def _text_render(
        text: str
) -> str:
    out_text = ""
    if text is not None:
        text = _newline_render(text)
        single_text_list = text.split('\n')
        out_text_list = []
        for single_text in single_text_list:
            text_piece = _single_text_render(single_text, max_characters=25)
            out_text_list.append(text_piece)
        
        for piece in out_text_list:
            out_text = out_text + piece + "\n"
        out_text = _newline_render(out_text)
    return out_text


def format_render(
        raw_dir:str,
) -> None:
    """Render the format for each individual report.
    
    Example
    -------
    >>> import raw
    >>> raw.format_render(r'./raw_reports')

    Parameters
    ----------
    raw_dir: str or path-like object
        directory of the raw reports

    .. note:
        Might need to update if the format of the monthly report changes as the 
        raw report for each student are picked the first 8 columns of the montly
        report.
    """
    report_paths = glob.glob(r'{}/*.xlsx'.format(raw_dir))
    print("Rendering formats...")
    with xw.App(visible=True, add_book=False) as app:
        for path in tqdm(report_paths):
            wb = app.books.open(path)
            sht = wb.sheets[0]
            range = utils._get_range(sht)
            cells = sht.range((2, 6), range)
            for cell in cells:
                text = cell.value
                cell.value = _text_render(text)
            sht.range('A1').column_width = 6
            sht.range('B1').column_width = 10
            sht.range('C1').column_width = 19
            sht.range('D1').column_width = 15
            sht.range('E1').column_width = 6
            sht.range('F1').column_width = 46
            sht.range('G1').column_width = 46
            sht.range('H1').column_width = 46
            sht.autofit()
            sht.range('A1:H1').color = [255, 192, 0]  # golden
            sht.range((2,1), range).interior_color = [255, 255, 255]  # white
            sht.range((2,1), range).font.bold = True  # bold text
            wb.save()
            wb.close()
    print("Format rendered.\n")



