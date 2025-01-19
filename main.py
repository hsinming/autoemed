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
import logging
from datetime import datetime
from pathlib import Path
from time import sleep
import openpyxl
from helium import *


EXCEL_PATH = Path("test2.xlsx")
EMEDICAL_URL = 'https://www.emedical.immi.gov.au/eMedUI/eMedical'
USER_ID = 'e26748'
PASSWORD = 'Year*2025'


# 設定日誌
log_date = datetime.now().strftime('%Y-%m-%d')
log_file = Path(f"{log_date}.txt")
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', handlers=[
    logging.FileHandler(log_file, mode='a', encoding='utf-8'),
    logging.StreamHandler()
])


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


def search(emed_no: str):
    click(RadioButton('Using Health Case Identifier'))
    write(emed_no, into=TextField('ID'))
    click(Button('Search'))
    wait_until(Text('Select:').exists)


def manage_case():
    click(Button('All'))
    click(Button('Manage Case'))
    wait_until(Text('Pre exam: Health case details').exists)


def process_australia(emed_no: str):
    try:
        search(emed_no)
        sleep(1)
        manage_case()

        if Text(below=Text('Health Case Status'), to_left_of=Text('Exam in Progress')).value == 'CURRENT':
            click(Text('502 Chest X-Ray Examination'))
            click(Text('Detailed radiology findings'))
            wait_until(Text('Detailed question').exists)

            for rb_normal in find_all(RadioButton('Normal')):
                if not rb_normal.is_selected():
                    click(rb_normal)

            if not RadioButton('Absent').is_selected():
                click(RadioButton('Absent'))

            if not RadioButton('No').is_selected():
                click(RadioButton('Absent'))

            # if not RadioButton('Normal', to_right_of=Text('1. Skeleton and soft tissue')).is_selected():
            #     click(RadioButton('Normal', to_right_of=Text('1. Skeleton and soft tissue')))
            #     click(RadioButton('Normal', to_right_of=Text('2. Cardiac shadow')))
            #     click(RadioButton('Normal', to_right_of=Text('3. Hilar and lymphatic glands')))
            #     click(RadioButton('Normal', to_right_of=Text('4. Hemidiaphragms and costophrenic angles')))
            #     click(RadioButton('Normal', to_right_of=Text('5. Lung fields')))
            #     click(RadioButton('Absent', to_right_of=Text('6. Evidence of Tuberculosis (TB)')))
            #     click(RadioButton('No', to_right_of=Text('7. Are there strong suspicions of active Tuberculosis (TB)?')))

            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Review exam details').exists)
            sleep(1)
            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Grading & Examiner Declaration').exists)

            if Button('Prepare for grading').exists() and Button('Prepare for grading').is_enabled():
                click(Button('Prepare for grading'))
                wait_until(Text('Provide Grading').exists)

            if not RadioButton(
                    'A - No evidence of active TB, or changes consistent with old or inactive TB, or changes suggestive of other significant diseases identified.').is_selected():
                click(RadioButton(
                    'A - No evidence of active TB, or changes consistent with old or inactive TB, or changes suggestive of other significant diseases identified.'))

            if not CheckBox(
                    'I declare that the chest X-ray examination report is a true and correct record of my findings.').is_checked():
                click(CheckBox(
                    'I declare that the chest X-ray examination report is a true and correct record of my findings.'))

            if Button('Submit Exam').exists() and Button('Submit Exam').is_enabled():
                click(Button('Submit Exam'))
                wait_until(Alert().exists)
                Alert().accept()
                wait_until(Text('Success').exists)

        click(Button('Close'))
        logging.info(f"成功處理 : {emed_no}")
        return True
    except Exception as e:
        logging.error(f"處理失敗 : {emed_no}, 錯誤: {e}")
        return False


def process_canada(emed_no: str):
    try:
        search(emed_no)
        sleep(1)
        manage_case()

        if Text(below=Text('Health Case Status'), to_left_of=Text('Exam in Progress')).value == 'CURRENT':
            click(Text('502 Chest X-Ray Examination'))
            click(Text('Detailed radiology findings'))
            wait_until(Text('Detailed question').exists)

            for rb_normal in find_all(RadioButton('Normal')):
                if not rb_normal.is_selected():
                    click(rb_normal)

            if not RadioButton('Absent').is_selected():
                click(RadioButton('Absent'))

            if not RadioButton('No').is_selected():
                click(RadioButton('Absent'))

            # if not RadioButton('Normal', to_right_of=Text('1. Skeleton and soft tissue')).is_selected():
            #     click(RadioButton('Normal', to_right_of=Text('1. Skeleton and soft tissue')))
            #     click(RadioButton('Normal', to_right_of=Text('2. Cardiac shadow')))
            #     click(RadioButton('Normal', to_right_of=Text('3. Hilar and lymphatic glands')))
            #     click(RadioButton('Normal', to_right_of=Text('4. Hemidiaphragms and costophrenic angles')))
            #     click(RadioButton('Normal', to_right_of=Text('5. Lung fields')))
            #     click(RadioButton('Absent', to_right_of=Text('6. Evidence of Tuberculosis (TB)')))
            #     click(RadioButton('No', to_right_of=Text('7. Are there strong suspicions of active Tuberculosis (TB)?')))

            click(Button('Next'))
            wait_until(Text('Special findings').exists)

            if not RadioButton('None of the following are present').is_selected():
                click(RadioButton('None of the following are present'))

            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Review exam details').exists)
            sleep(1)
            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Grading & Examiner Declaration').exists)

            if Button('Prepare for grading').exists() and Button('Prepare for grading').is_enabled():
                click(Button('Prepare for grading'))
                wait_until(Text('Provide Grading').exists)

            if not RadioButton(
                    'A - No evidence of active TB, or changes consistent with old or inactive TB, or changes suggestive of other significant diseases identified.').is_selected():
                click(RadioButton(
                    'A - No evidence of active TB, or changes consistent with old or inactive TB, or changes suggestive of other significant diseases identified.'))

            if not CheckBox(
                    'I declare that the chest X-ray examination report is a true and correct record of my findings.').is_checked():
                click(CheckBox(
                    'I declare that the chest X-ray examination report is a true and correct record of my findings.'))

            if Button('Submit Exam').exists() and Button('Submit Exam').is_enabled():
                click(Button('Submit Exam'))
                wait_until(Alert().exists)
                Alert().accept()
                wait_until(Text('Success').exists)

        click(Button('Close'))
        logging.info(f"成功處理 : {emed_no}")
        return True
    except Exception as e:
        logging.error(f"處理失敗 : {emed_no}, 錯誤: {e}")
        return False


def process_new_zealand(emed_no: str):
    try:
        search(emed_no)
        sleep(1)
        manage_case()

        if Text(below=Text('Health Case Status'), to_left_of=Text('Exam in Progress')).value == 'CURRENT':
            click(Text('502 Chest X-Ray Examination'))
            click(Text('Detailed radiology findings'))
            wait_until(Text('Detailed question').exists)

            for rb_normal in find_all(RadioButton('Normal')):
                if not rb_normal.is_selected():
                    click(rb_normal)

            if not RadioButton('Absent').is_selected():
                click(RadioButton('Absent'))

            if not RadioButton('No').is_selected():
                click(RadioButton('Absent'))

            # if not RadioButton('Normal', to_right_of=Text('1. Skeleton and soft tissue')).is_selected():
            #     click(RadioButton('Normal', to_right_of=Text('1. Skeleton and soft tissue')))
            #     click(RadioButton('Normal', to_right_of=Text('2. Cardiac shadow')))
            #     click(RadioButton('Normal', to_right_of=Text('3. Hilar and lymphatic glands')))
            #     click(RadioButton('Normal', to_right_of=Text('4. Hemidiaphragms and costophrenic angles')))
            #     click(RadioButton('Normal', to_right_of=Text('5. Lung fields')))
            #     click(RadioButton('Absent', to_right_of=Text('6. Evidence of Tuberculosis (TB)')))
            #     click(RadioButton('No', to_right_of=Text('7. Are there strong suspicions of active Tuberculosis (TB)?')))

            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Review exam details').exists)
            sleep(1)
            click(Button('Next'))
            wait_until(Text('502 Chest X-Ray Examination: Grading & Examiner Declaration').exists)

            if Button('Prepare for grading').exists() and Button('Prepare for grading').is_enabled():
                click(Button('Prepare for grading'))
                wait_until(Text('Provide Grading').exists)

            if not RadioButton(
                    'A - No evidence of active TB, or changes consistent with old or inactive TB, or changes suggestive of other significant diseases identified.').is_selected():
                click(RadioButton(
                    'A - No evidence of active TB, or changes consistent with old or inactive TB, or changes suggestive of other significant diseases identified.'))

            if not CheckBox(
                    'I declare that the chest X-ray examination report is a true and correct record of my findings.').is_checked():
                click(CheckBox(
                    'I declare that the chest X-ray examination report is a true and correct record of my findings.'))

            if Button('Submit Exam').exists() and Button('Submit Exam').is_enabled():
                click(Button('Submit Exam'))
                wait_until(Alert().exists)
                Alert().accept()
                wait_until(Text('Success').exists)

        click(Button('Close'))
        logging.info(f"成功處理 : {emed_no}")
        return True

    except Exception as e:
        logging.error(f"處理失敗 : {emed_no}, 錯誤: {e}")
        return False


def process_united_states(emed_no: str):
    try:
        search(emed_no)
        sleep(1)
        manage_case()

        if Text(below=Text('Health Case Status'), to_left_of=Text('Exam in Progress')).value == 'CURRENT':
            click(Text('502 Chest X-Ray Examination'))
            click(Text('Findings'))
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

            if not CheckBox(
                    'I declare that the chest X-ray examination report is a true and correct record of my findings.').is_checked():
                click(CheckBox(
                    'I declare that the chest X-ray examination report is a true and correct record of my findings.'))

            if Button('Submit Exam').exists() and Button('Submit Exam').is_enabled():
                click(Button('Submit Exam'))
                wait_until(Alert().exists)
                Alert().accept()
                wait_until(Text('Success').exists)

        click(Button('Close'))
        logging.info(f"成功處理 : {emed_no}")
        return True

    except Exception as e:
        logging.error(f"處理失敗 : {emed_no}, 錯誤: {e}")
        return False


if __name__ == "__main__":
    # Login to eMedical website
    start_chrome(EMEDICAL_URL)
    write(USER_ID, into=TextField('User id'))
    write(PASSWORD, into=TextField('Password'))
    click(Button('Logon'))
    wait_until(Text('Case search').exists, timeout_secs=10)

    emedical_number_list = extract_emedical_no(EXCEL_PATH)
    logging.info(f"讀取 {len(emedical_number_list)} 個 eMedical No.")
    # print(emedical_number_list)

    success_list = []
    failure_list = []

    for emed_no in emedical_number_list[:1]:
        logging.info(f'現在處理: {emed_no}')

        if emed_no.startswith(('HAP', 'TRN')):
            success = process_australia(emed_no)
            # pass
        elif emed_no.startswith(('NZER', 'NZHR')):
            success = process_new_zealand(emed_no)
            # pass
        elif emed_no.startswith(('IME', 'UMI', 'UCI')):
            success = process_canada(emed_no)
            # pass
        elif emed_no.startswith('CEAC'):
            success = process_united_states(emed_no)
            # pass
        else:
            logging.warning(f"未知的 eMedical No. 類型: {emed_no}")
            success = False

        if success:
            success_list.append(emed_no)
        else:
            failure_list.append(emed_no)

    logging.info("處理完成！")
    logging.info(f"成功: {len(success_list)}, 失敗: {len(failure_list)}")

