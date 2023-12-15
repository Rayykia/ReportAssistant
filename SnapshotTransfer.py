r"""
Grab snapshot of the raw report and transfer it to the tamplate.

Example
-------
::
    >>> import SnapshotTransfer

        Grab images form raw reports.

    >>> SnapshotTransfer.grab_image(r'../../raw_reports', r'../../images')

        Insert images to the template.xlsx

    >>> improt xlwings as xs

    >>> with xw.App(visible=True, add_books=False) as app:
            sht = app.books.open("template.xlsx")[0]
            SnapshotTransfer.insert_pic(sht, student)  # student: png file of the student

Author: Rayykia

Date: 12-7-2023
"""

import xlwings as xw
import os
import glob

from PIL import Image
from tqdm import tqdm
from typing import Union
import utils


def grab_image(
        dir: str,
        o_dir: str, 
        sheet: Union[int, str] = 0
) -> None:
    """Grab the images of each raw report.
    
    Example
    -------
    ::
        >>> from utils import SnapshotTransfer

        >>> SnapshotTransfer.grab_image(r'../../raw_reports', r'../../images')

    Parameters
    ----------
    dir: str or path-like object
        the directory of `raw reports` (processed by package `raw`)
    o_dir: str or path-like object
        the output directory
    sheet: int or str
        the sheet used to grab image, `index` or `name`

    """
    if not os.path.isdir(o_dir):
        os.makedirs(o_dir)

    student_paths = glob.glob(r"{}/*.xlsx".format(dir))

    print("Grabbing images...")
    with xw.App(visible=False, add_book=False) as app:
        for path in tqdm(student_paths):
            wb = app.books.open(path)
            sht = wb.sheets[sheet]

            student_name = sht.range((2, 1)).value
            
            cells = sht.range((1, 1), utils._get_range(sht))
            cells.to_png(r"{}/{}.png".format(o_dir, student_name))
            wb.close()
    print("Image grabbed.\n")


def insert_pic(
        sht: xw.Sheet, 
        path: str
) -> None:
    """Insert the picture to the tempalte excel.

    Parameters
    ----------
    sht: xlwings.Sheet
        the template sheet
    path: str or path-like object
        the path of the picture
    """
    img = Image.open(path)
    width = sht.range('b6').width
    ratio = width/img.size[0]
    height = img.size[1]*ratio
    block_height = height+3
    if block_height <= 409:
        sht.range('a6').api.RowHeight = block_height
    else:
        res = block_height - 409
        n = int(res // 409 + 1)
        remainder = res - (n-1)*409
        for i in range(n):
            sht.api.Rows(6).Insert()
            sht.range('a6').api.RowHeight = 409
        sht.range('a6').api.RowHeight = remainder
        sht.range((6, 1), (n+6, 1)).merge()
        sht.range((6, 2), (n+6, 2)).merge()
    sht.pictures.add(path, left=sht.range("B6").left, top=sht.range("B6").top, width=width, height=height, name='report')

