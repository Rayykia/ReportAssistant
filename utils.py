r"""
Utils for ReportAssistant.

Author: Rayykia

Date: 12-7-2023
"""

import xlwings as xw
import pandas as pd

from typing import Tuple

def _get_range(
        sheet: xw.Sheet
) -> str:
    """Get the range of a report.

    Parameters
    ----------
    sheet: xlwings.Sheet
        the sheet for range grabbing

    .. note:
        Might need to update if the format of the monthly report changes as the 
        raw report for each student are picked the first 8 columns of the montly
        report.
    """
    n_rows = sheet.range('A1').end('down').row
    return (n_rows, 8)


def get_names(
        data: pd.DataFrame
) -> Tuple[list, list]:
    """Get the name list of the students from the list of report's paths.

    Example
    -------
    >>> import utils
    >>> stu_list, name_list = get_names(report_dataframe)

    Parameters
    ----------
    data: pandas.DataFrame
        data contains `学生`
    """
    stu_list = list(data["学生"].unique())
    return stu_list

