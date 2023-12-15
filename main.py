"""New Channel Report Assistant.

Copyright (C) 2023, Rayykia.
All rights reserved.

License: BSD License (see LICENSE.txt for details)

Author: Rayykia

.. version: 1.0

.. notes:
    Collects information for the following data format:
    |学生|--|上课日期|--|上课时段|--|课程|--|教师|--|复习检查|--|课堂内容|--|学生表现|.
    This version of ReportAssistant is no longer available if the above format is altered.

"""

import xlwings as xw
import pandas as pd
import glob
import sys
import os
import re
import calendar


import raw
import SnapshotTransfer
import remark
from tqdm import tqdm
from typing import Tuple


def construct_directory(
        current_dir: str,
        year: int,
        month: int,
        supervisor: str,
) -> Tuple[str, ...]:
    """Construct the working directories.
    ::
    current_dir/monthly_reports/ --------[year1]_[month1]_[supervisor] -------- final_reports
                                    |                                      |
                                    |                                      |--- images
                                    |                                      |
                                    |                                      ---- raw_reports
    construct when generating       |
    another month's report ======>  |----[year2]_[month2]_[supervisor] -------- final_reports
                                                                           |
                                                                           |--- images
                                                                           |
                                                                           ---- raw_reports
    Properties
    ----------
    current_dir: str or path-like object
        the current directory
    year: int
        the year of the report
    month: int
        the month of the report
    supervisor: str
        the name of the `academic supervisor`

    Returns
    -------
    subfile_paths: tuple
        path of `final_reports`, `raw_reports`, `images`
        (final_path, raw_path, images_path)
    """
    bin = "monthly_reports"
    if not os.path.isdir(bin):
        os.mkdir(bin)

    base = "monthly_reports/{}_{}_{}".format(year, month, supervisor)
    subfiles = ["final_reports", "raw_reports", "images"]

    subfile_paths = ()
    for file in subfiles:
        subfile_paths = subfile_paths.__add__((os.path.join(current_dir, base, file),))
    
    for path in [base, *subfile_paths]:
        if not os.path.isdir(path):
            os.mkdir(path)
    
    return subfile_paths




def main() -> None:
    """
    The main function of the program. 
    =================================

    It generate monthly reports with the following pipeline:
    1. Generate raw reports.
    2. Grab images from raw reports of each student.
    3. Generate remarks for each studemt accordingly.
    4. Fill the report template for each student.

    """
    data_path = str(input("Select file (.xls or .xlsx): "))
    data = pd.read_excel(data_path)
    data = data.iloc[:,:8]

    # Set up directory paths
    current_dir =os.path.dirname(os.path.realpath(sys.argv[0]))
    template_path = os.path.join(current_dir, "ReportAssistant_bin/template.xlsx")
    corpus_path = os.path.join(current_dir, "ReportAssistant_bin/remarks.txt")
    info_path = os.path.join(current_dir, "ReportAssistant_bin/supervisor_info.txt")

    year = data["上课日期"][0].split("-")[0]
    month = data["上课日期"][0].split("-")[1]
    print("\nGenerating reports for {} - {}.".format(str(month), str(year)), end="\n\n")
    last_day = str(calendar.monthrange(int(year), int(month))[-1])

    with open(info_path, "r", encoding="utf-8") as f:
        supervisors_info = f.read()
    supervisors = supervisors_info.replace(" ", "").split("/")  


    
    final_path, raw_path, images_path = construct_directory(
        current_dir, year, month, supervisors[1]
    )

    raw.split_book(data, raw_path)

    


    remark_list, remark_info = remark.get_info(data)

    remark_names = [re.sub('[a-zA-Z]','',x)[-2:] for x in remark_list]  # the last 2 characters of each name

    raw.format_render(raw_path)
    SnapshotTransfer.grab_image(raw_path, images_path)

    path_list= [os.path.abspath(x) for x in glob.glob(images_path+"\\*.png")]

    

 
    # Generate reports for each student
    print("Generating repoers...")
    with xw.App(visible=True, add_book=False) as app:
        for student in tqdm(path_list):
            student_name = student.split("\\")[-1].split(".")[0]
            wb = app.books.open(template_path)
            sht = wb.sheets[0]
            
            sht.range("A1").value = sht.range("A1").value.replace("xxx", student_name)
            sht.range("B2").value = student_name
            date_text = sht.range("B3").value
            date_text = date_text.replace("xxxx", year).replace("xx", month).replace("x", last_day)
            sht.range("B3").value = date_text
            sht.range("B4").value = supervisors[0]
            sht.range("B5").value = supervisors[1]

            if student_name in remark_list:
                index = remark_list.index(student_name)
                remark_text = remark.generate_reamrk(corpus_path, remark_names[index], remark_info[index])
                sht.range("B7").value = remark_text

            SnapshotTransfer.insert_pic(sht, student)
            wb.save(
                os.path.join(final_path, "新航道万象城校区{}{}月学习总结.xlsx".format(student_name, month))
            )
            wb.close()
            del wb, sht
    print("Done!")



if __name__ == '__main__':
    print("\nNew Channel Report Assistant. [Version 1.0]")
    print("Copyright (c) 2023, Rayykia.")
    print("All rights reserved.", end="\n\n")

    main()