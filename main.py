#!/usr/bin/env python3
# -*- coding:utf-8 -*-
"""
@author: Hsin-ming Chen
@license: GPL
@file: main.py
@time: 2025/01/15
@contact: hsinming.chen@gmail.com
@software: PyCharm
"""
from pathlib import Path
from time import sleep
import openpyxl
from helium import *


EXCEL_PATH = Path("test2.xlsx")
EMEDICAL_URL = 'https://www.emedical.immi.gov.au/eMedUI/eMedical'
USER_ID = 'e26748'
PASSWORD = 'Year*2025'


def is_black_font(cell):
    return cell.font is None or cell.font.color is None or cell.font.color.rgb in ["FF000000", None]


def is_no_fill(cell):
    return cell.fill is None or cell.fill.fgColor is None or cell.fill.fgColor.rgb in ["00000000", "FFFFFFFF", None]


def extract_emedical_no(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    all_emedical_nos = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.strip() == "eMedical No.":
                    # 獲取此列下方的所有值
                    col_idx = cell.column
                    for r in range(cell.row + 1, ws.max_row + 1):
                        target_cell = ws.cell(row=r, column=col_idx)
                        if target_cell.value and is_black_font(target_cell) and is_no_fill(target_cell):
                            all_emedical_nos.append(target_cell.value)

    wb.close()
    return all_emedical_nos


def process_australia(emed_no: str):
    click(RadioButton('Using Health Case Identifier'))
    write(emed_no, TextField('ID'))
    click(Button('Search'))

    if not Text('Your search returned no results.').exists():
        click(S("#caseSearch-searchResults_0", above=Button('Manage Case')))    # the first match ID
        click(Button('Manage Case'))
        wait_until(Text('Pre exam: Health case details').exists)

        if Text(below=Text('Health Case Status'), to_left_of=Text('Exam in Progress')).value == 'CURRENT':
            click(Text('502 Chest X-Ray Examination', below=Text('All Exams')))
            click(Text('Detailed radiology findings', below=Text('502 Chest X-Ray Examination')))
            wait_until(Text('502 Chest X-Ray Examination: Detailed radiology findings').exists)

            if not RadioButton('Normal', to_right_of=Text('1. Skeleton and soft tissue')).is_selected():
                click(RadioButton('Normal', to_right_of=Text('1. Skeleton and soft tissue')))
                click(RadioButton('Normal', to_right_of=Text('2. Cardiac shadow')))
                click(RadioButton('Normal', to_right_of=Text('3. Hilar and lymphatic glands')))
                click(RadioButton('Normal', to_right_of=Text('4. Hemidiaphragms and costophrenic angles')))
                click(RadioButton('Normal', to_right_of=Text('5. Lung fields')))
                click(RadioButton('Absent', to_right_of=Text('6. Evidence of Tuberculosis (TB)')))
                click(RadioButton('No', to_right_of=Text('7. Are there strong suspicions of active Tuberculosis (TB)?')))
            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Review exam details').exists)
            sleep(1)
            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Grading & Examiner Declaration').exists)

            if Button('Prepare for grading').exists() and Button('Prepare for grading').is_enabled():
                click(Button('Prepare for grading'))
                wait_until(Text('Provide Grading').exists)

            if Button('Submit Exam').exists() and Button('Submit Exam').is_enabled():
                click(RadioButton(
                    'A - No evidence of active TB, or changes consistent with old or inactive TB, or changes suggestive of other significant diseases identified.'))
                click(CheckBox(
                    'I declare that the chest X-ray examination report is a true and correct record of my findings.'))
                click(Button('Submit Exam'))
                Alert().accept()
                wait_until(Text('Success').exists)

        click(Button('Close'))


def process_canada(emed_no: str):
    click(RadioButton('Using Health Case Identifier'))
    write(emed_no, TextField('ID'))
    click(Button('Search'))

    if not Text('Your search returned no results.').exists():
        click(S("#caseSearch-searchResults_0", above=Button('Manage Case')))    # the first match ID
        click(Button('Manage Case'))
        wait_until(Text('Pre exam: Health case details').exists)

        if Text(below=Text('Health Case Status'), to_left_of=Text('Exam in Progress')).value == 'CURRENT':
            click(Text('502 Chest X-Ray Examination', below=Text('All Exams')))
            click(Text('Detailed radiology findings', below=Text('502 Chest X-Ray Examination')))
            wait_until(Text('502 Chest X-Ray Examination: Detailed radiology findings').exists)

            if not RadioButton('Normal', to_right_of=Text('1. Skeleton and soft tissue')).is_selected():
                click(RadioButton('Normal', to_right_of=Text('1. Skeleton and soft tissue')))
                click(RadioButton('Normal', to_right_of=Text('2. Cardiac shadow')))
                click(RadioButton('Normal', to_right_of=Text('3. Hilar and lymphatic glands')))
                click(RadioButton('Normal', to_right_of=Text('4. Hemidiaphragms and costophrenic angles')))
                click(RadioButton('Normal', to_right_of=Text('5. Lung fields')))
                click(RadioButton('Absent', to_right_of=Text('6. Evidence of Tuberculosis (TB)')))
                click(RadioButton('No', to_right_of=Text('7. Are there strong suspicions of active Tuberculosis (TB)?')))
            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Special findings').exists)

            click(RadioButton('None of the following are present'))
            click(Button('Next'))

            wait_until(Text('502 Chest X-Ray Examination: Review exam details').exists)
            sleep(1)
            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Grading & Examiner Declaration').exists)

            if Button('Prepare for grading').exists() and Button('Prepare for grading').is_enabled():
                click(Button('Prepare for grading'))
                wait_until(Text('Provide Grading').exists)

            if Button('Submit Exam').exists() and Button('Submit Exam').is_enabled():
                click(RadioButton(
                    'A - No evidence of active TB, or changes consistent with old or inactive TB, or changes suggestive of other significant diseases identified.'))
                click(CheckBox(
                    'I declare that the chest X-ray examination report is a true and correct record of my findings.'))
                click(Button('Submit Exam'))
                Alert().accept()
                wait_until(Text('Success').exists)

        click(Button('Close'))


def process_new_zealand(emed_no: str):
    click(RadioButton('Using Health Case Identifier'))
    write(emed_no, TextField('ID'))
    click(Button('Search'))

    if not Text('Your search returned no results.').exists():
        click(S("#caseSearch-searchResults_0", above=Button('Manage Case')))  # the first match ID
        click(Button('Manage Case'))
        wait_until(Text('Pre exam: Health case details').exists)

        if Text(below=Text('Health Case Status'), to_left_of=Text('Exam in Progress')).value == 'CURRENT':
            click(Text('502 Chest X-Ray Examination', below=Text('All Exams')))
            click(Text('Detailed radiology findings', below=Text('502 Chest X-Ray Examination')))
            wait_until(Text('502 Chest X-Ray Examination: Detailed radiology findings').exists)

            if not RadioButton('Normal', to_right_of=Text('1. Skeleton and soft tissue')).is_selected():
                click(RadioButton('Normal', to_right_of=Text('1. Skeleton and soft tissue')))
                click(RadioButton('Normal', to_right_of=Text('2. Cardiac shadow')))
                click(RadioButton('Normal', to_right_of=Text('3. Hilar and lymphatic glands')))
                click(RadioButton('Normal', to_right_of=Text('4. Hemidiaphragms and costophrenic angles')))
                click(RadioButton('Normal', to_right_of=Text('5. Lung fields')))
                click(RadioButton('Absent', to_right_of=Text('6. Evidence of Tuberculosis (TB)')))
                click(RadioButton('No', to_right_of=Text('7. Are there strong suspicions of active Tuberculosis (TB)?')))
            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Review exam details').exists)
            sleep(1)
            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Grading & Examiner Declaration').exists)

            if Button('Prepare for grading').exists() and Button('Prepare for grading').is_enabled():
                click(Button('Prepare for grading'))
                wait_until(Text('Provide Grading').exists)

            if Button('Submit Exam').exists() and Button('Submit Exam').is_enabled():
                click(RadioButton(
                    'A - No evidence of active TB, or changes consistent with old or inactive TB, or changes suggestive of other significant diseases identified.'))
                click(CheckBox(
                    'I declare that the chest X-ray examination report is a true and correct record of my findings.'))
                click(Button('Submit Exam'))
                Alert().accept()
                wait_until(Text('Success').exists)

        click(Button('Close'))


def process_united_states(emed_no: str):
    click(RadioButton('Using Health Case Identifier'))
    write(emed_no, TextField('ID'))
    click(Button('Search'))

    if not Text('Your search returned no results.').exists():
        click(S("#caseSearch-searchResults_0", above=Button('Manage Case')))    # the first match ID
        click(Button('Manage Case'))
        wait_until(Text('Pre exam: Health case details').exists)

        if Text(below=Text('Health Case Status'), to_left_of=Text('Exam in Progress')).value == 'CURRENT':
            click(Text('502 Chest X-Ray Examination', below=Text('All Exams')))
            click(Text('Findings', below=Text('502 Chest X-Ray Examination')))
            wait_until(Text('502 Chest X-Ray Examination: Findings').exists)

            if not RadioButton('Normal', to_right_of=Text('Findings')).is_selected():
                click(RadioButton('Normal', to_right_of=Text('Findings')))

            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Review exam details').exists)
            sleep(1)
            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Examiner Declaration').exists)

            if Button('Prepare for declaration').exists() and Button('Prepare for declaration').is_enabled():
                click(Button('Prepare for declaration'))
                wait_until(Text('Examiner declaration').exists)

            if Button('Submit Exam').exists() and Button('Submit Exam').is_enabled():
                click(CheckBox(
                    'I declare that the chest X-ray examination report is a true and correct record of my findings.'))
                click(Button('Submit Exam'))
                Alert().accept()
                wait_until(Text('Success').exists)

        click(Button('Close'))


if __name__ == "__main__":
    # Start Chrome and navigate to the eMedical login page
    driver = start_chrome(EMEDICAL_URL)
    write(USER_ID, into=TextField('User id'))
    write(PASSWORD, into=TextField('Password'))
    click(Button('Logon'))
    wait_until(Text('Case search').exists, timeout_secs=10)

    emedical_number_list = extract_emedical_no(EXCEL_PATH)
    print(emedical_number_list)

    for emed_no in emedical_number_list[18:19]:
        print(emed_no)

        if emed_no.startswith(('HAP', 'TRN')):
            process_australia(emed_no)
            # pass
        elif emed_no.startswith(('NZER', 'NZHR')):
            process_new_zealand(emed_no)
            # pass
        elif emed_no.startswith(('IME', 'UMI', 'UCI')):
            process_canada(emed_no)
            # pass
        elif emed_no.startswith('CEAC'):
            process_united_states(emed_no)
            # pass
        else:
            pass

