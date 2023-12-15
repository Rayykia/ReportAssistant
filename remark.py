"""
Generate remarks for each student.

Example
-------
::
    >>> import remark

    >>> remark_list, remark_info = remark.get_info(data)
    
    >>> if student_name in remark_list:
            index = remark_list.index(student_name)
            remark_text = remark.generate_reamrk(corpus_path, remark_names[index], remark_info[index])
            sht.range("B7").value = remark_text

Author: Rayykia

Date: 12-8-2023

.. notes:
    The program determins whether the student finished his or her homework by checking 
    whether there's `没` or `未` in the `复习检查` column. Might not accurate, please double check.

"""

import pandas as pd

import utils

from typing import Tuple


def get_info(
        data: pd.DataFrame
) -> Tuple[list, list]:
    """Get the info of each student.

    Determines whether a student needs remark.

    Determins whether the student finished his or her homework by checking 
    whether there's `没` or `未` in the `复习检查` column.

    Example
    -------
    ::
        >>> import reamrk

            data: dataframe that contains `学生`, `课程` and `复习检查`
        >>> remark_list, remark_info = remark.get_info(data)

    parameters
    ----------
    data: pandas.DataFrame
        data contains `学生`, `课程` and `复习检查`

    Returns
    -------
    remark_list: list
        the names of studnets taking the subject in ["留学预备","托福","雅思","sat"]
    remark_info: list
        each element is a list, [`name`, `subject`, `0 or 1`]
        `0` means the assignmetn is finished, `1` indicats the assignment is not finished

    .. notes:
        The check of the assignmemt might not be accurate, please double check.

    """
    not_finished = ["没","未"]
    check_list_1 = data.iloc[:, 5].str.contains(not_finished[0])  # `复习检查`
    check_list_2 = data.iloc[:, 5].str.contains(not_finished[1])
    chekc_list = [bool(x) for x in check_list_1+check_list_2]
    unfinished_list = data.iloc[:, 0][chekc_list].unique()  # `学生`

    subject_list = ["留学预备","托福","雅思","sat"]
    students = utils.get_names(data)
    student_subjects = []
    for subject in subject_list:
        student_subjects.append(list(data.iloc[:, 0][data.iloc[:, 3].str.contains(subject)].unique()))  # `学生`, `课程`

    remark_info = []
    for i, subject in enumerate(subject_list):
        for student in students:
            if student in student_subjects[i]:
                remark_info.append([student, subject])
    remark_list = [x[0] for x in remark_info]
    for j, student in enumerate(remark_list):
        if student in unfinished_list:
            remark_info[j].append(1)
        else:
            remark_info[j].append(0)
    return remark_list, remark_info




def generate_reamrk(
        corpus_path: str,
        name: str,
        remark_info: list
) -> str:
    """Generate remark for a student.
    ["留学预备","托福","雅思","sat"]
    
    For the first part of the reamrk is generated according to the remark_info:
    `first name`, `subject` and `whether the student finished the assignment`.
    , first name and subject is filled. 

    The suggestion for the study plans is generated according to the 
    the student's subject.

    Example
    -------
    ::
        >>> import remark

            corpus: txt file that contains remarks for each subject
        >>> remark_text = remark.generate_reamrk(corpus_path, remark_names, remark_info)


    Parameters
    ----------
    name: str
        the first name of the student
    remark_info: list
        information about [`name`, `subject`, `assignment`]

    Returns
    -------
    remark: str
        the remark for the student
    """
    subject = remark_info[1]

    with open(corpus_path,'r', encoding="utf-8") as f:
        remark_base = f.read()
    remark_base = remark_base.split("%")

    
    text = [remark_base[0].replace("[name]", name).replace("[subject]", subject)]

    if remark_info[-1] == 0:
        if subject == "留学预备":
            text.append(remark_base[2])
        if subject in ["托福","雅思"]:
            text.append(remark_base[3])
        if subject == "sat":
            text.append(remark_base[1])
    else:
        text.append(remark_base[4])

    text.append("\n")
    if subject == "留学预备":
        text.append(remark_base[5])
    elif subject == "托福":
        text.append(remark_base[6])
    elif subject == "雅思":
        text.append(remark_base[7])
    if subject == "sat":
        text.append(remark_base[8])

    remark = "".join(text)

    return remark