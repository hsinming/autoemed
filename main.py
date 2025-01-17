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
from typing import List
from pathlib import Path
from time import sleep
from openpyxl import load_workbook
import pandas as pd
from helium import *


EXCEL_PATH = Path("test2.xlsx")
EMEDICAL_URL = 'https://www.emedical.immi.gov.au/eMedUI/eMedical'
USER_ID = 'e26748'
PASSWORD = 'Year*2025'


def login_to_emedical(user_id: str, password: str):
    """
    Opens Chrome browser, navigates to the eMedical login page, and logs in with the provided credentials.

    Parameters:
    user_id (str): The user ID for login.
    password (str): The password for login.
    """
    # Start Chrome and navigate to the eMedical login page
    start_chrome('https://www.emedical.immi.gov.au/eMedUI/eMedical')

    # Enter the user ID and password
    write(user_id, into='User id')
    write(password, into='Password')

    # Click the Logon button
    click('Logon')

    # Optional: Add a wait or check to confirm successful login
    wait_until(Text('Logout').exists, timeout_secs=10)


def extract_all_eMedical_no_black_text(file_path) -> List[str]:
    # Load workbook
    workbook = load_workbook(file_path, data_only=True)
    eMedical_no_list = []

    # Iterate over all sheets
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        # Find the header "eMedical No."
        for col in sheet.iter_cols():
            for cell in col:
                if cell.value == "eMedical No.":
                    # Start collecting values from the next row below the header
                    row_idx = cell.row + 1
                    while row_idx <= sheet.max_row:
                        value_cell = sheet.cell(row=row_idx, column=cell.column)
                        # Check if the cell has text, no fill color, and black text color
                        if value_cell.value:
                            fill_color = value_cell.fill.start_color.index
                            font_color = value_cell.font.color.rgb if value_cell.font.color else None
                            if fill_color == '00000000' and (font_color == '00000000' or font_color is None):
                                eMedical_no_list.append(value_cell.value)
                        row_idx += 1

    return eMedical_no_list


def extract_emedical_numbers(file_path) -> List:
    """
    Extracts all eMedical No. from an Excel file and returns them as a list.

    Parameters:
    file_path (str): The path to the Excel file.

    Returns:
    List[str]: A list of all eMedical No. found in the file.
    """
    # Load the Excel file
    excel_content = pd.ExcelFile(file_path)
    all_emedical_numbers = []

    # Iterate through each sheet to extract eMedical No.
    for sheet in excel_content.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet)
        # Extract columns that potentially contain eMedical No.
        for col in df.columns:
            if df[col].dtype == object:  # Check if the column contains strings
                id_type = r'(HAP \d+|TRN \w+|NZER \w+|NZHR \w+|IME \w+|UMI \w+|UCI \w+|CEACBC \w+)'
                emedical_numbers = df[col].dropna().str.extract(id_type)[0]
                emedical_numbers = emedical_numbers.dropna().tolist()
                all_emedical_numbers.extend(emedical_numbers)

    return all_emedical_numbers


def process_australia(emed_no: str):
    click(RadioButton('Using Health Case Identifier'))
    write(emed_no, TextField('ID'))
    click(Button('Search'))
    if not Text('Your search returned no results.').exists():
        click(Text('All'))
        click(Button('Manage Case'))
        wait_until(Text('Pre exam: Health case details').exists)
        click(Button('Next'))
        wait_until(Text('Pre exam: Manage Photo').exists)
        sleep(3)  # wait loading photo
        click(Button('Next'))
        wait_until(Text('Pre exam: Confirm identity').exists)
        click(Button('Next'))
        wait_until(Text('All Exams: All exams summary').exists)
        click(Button('Close'))


def process_canada(emed_no: str):
    click(RadioButton('Using Health Case Identifier'))
    write(emed_no, 'ID')
    click(Button('Search'))
    if not Text('Your search returned no results.').exists():
        click(Text('All'))
        click(Button('Manage Case'))
        wait_until(Text('Pre exam: Health case details').exists)
        click(Button('Next'))
        wait_until(Text('Pre exam: Manage Photo').exists)
        sleep(3)  # wait loading photo
        click(Button('Next'))
        wait_until(Text('Pre exam: Confirm identity').exists)
        click(Button('Next'))
        wait_until(Text('All Exams: All exams summary').exists)
        click(Button('Close'))


def process_new_zealand(emed_no: str):
    pass


def process_united_states(emed_no: str):
    click(RadioButton('Using Health Case Identifier'))
    write(emed_no, 'ID')
    click(Button('Search'))
    if not Text('Your search returned no results.').exists():
        click(Text('All'))
        click(Button('Manage Case'))
        wait_until(Text('Pre exam: Health case details').exists)
        click(Button('Next'))
        wait_until(Text('Pre exam: Manage Photo').exists)
        sleep(3)  # wait loading photo
        click(Button('Next'))
        wait_until(Text('Pre exam: Confirm identity').exists)
        click(Button('Next'))
        wait_until(Text('All Exams: All exams summary').exists)
        click(Button('Close'))


if __name__ == "__main__":
    # Start Chrome and navigate to the eMedical login page
    driver = start_chrome(EMEDICAL_URL)
    write(USER_ID, into=TextField('User id'))
    write(PASSWORD, into=TextField('Password'))
    click(Button('Logon'))
    wait_until(Text('Case search').exists, timeout_secs=10)

    emedical_number_list = extract_all_eMedical_no_black_text(EXCEL_PATH)
    print(emedical_number_list)

    for emed_no in emedical_number_list[:1]:
        if emed_no.startswith(('HAP', 'TRN')):
            # process_australia(emed_no)
            pass
        elif emed_no.startswith(('NZER', 'NZHR')):
            # process_new_zealand(emed_no)
            pass
        elif emed_no.startswith(('IME', 'UMI', 'UCI')):
            # process_canada(emed_no)
            pass
        elif emed_no.startswith('CEAC'):
            # process_united_states(emed_no)
            pass
        else:
            pass

