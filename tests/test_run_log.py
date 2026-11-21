"""Run Log tests through EmedicalWorkflowManager.start_workflow (ADR-0002: fakes only, no Live Portal)."""
import json
import threading
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.styles import Font


class FakeExcel:
    def __init__(self, ids):
        self.ids = ids

    def extract_emedical_no(self, path):
        return list(self.ids)


class FakeAutomator:
    def __init__(self, main_module, results=None, login_ok=True, on_run=None):
        self.main = main_module
        self.results = results or {}
        self.login_ok = login_ok
        self.on_run = on_run

    def login(self):
        return self.login_ok

    def run_cxr_exam(self, emed_no, country):
        if self.on_run:
            self.on_run(emed_no)
        return self.results.get(emed_no, self.main.CxrResult(True, 'Close'))


@pytest.fixture
def run(main_module, tmp_path, monkeypatch):
    monkeypatch.setattr(main_module, 'stop_event', threading.Event())

    def _run(excel_processor, automator, excel_name='cases.xlsx'):
        excel = tmp_path / excel_name
        if not excel.exists():
            excel.write_bytes(b'')
        manager = main_module.EmedicalWorkflowManager(excel_processor, automator)
        manager.set_gui_callbacks(*[lambda *a, **k: None] * 6)
        manager.start_workflow(str(excel), close_browser=False)

    return _run


def read_run_log(tmp_path):
    files = list((tmp_path / 'run_logs').glob('*.jsonl'))
    assert len(files) == 1
    return [json.loads(line) for line in files[0].read_text(encoding='utf-8').splitlines()]


def test_all_success(main_module, run, tmp_path):
    run(FakeExcel(['HAP1', 'CEAC 2']), FakeAutomator(main_module))
    records = read_run_log(tmp_path)
    assert [r['kind'] for r in records] == ['run_start', 'case', 'case', 'run_end']
    start, end = records[0], records[-1]
    assert start['version'] == main_module.VERSION
    assert start['excel_file'] == 'cases.xlsx'
    assert start['case_count'] == 2
    assert start['start_time']
    assert [r['emed_no'] for r in records[1:3]] == ['HAP1', 'CEAC 2']
    assert [r['country'] for r in records[1:3]] == ['澳大利亞', '美國']
    for case in records[1:3]:
        assert case['outcome'] == 'success'
        assert case['step'] == 'Close'
        assert case['duration'] >= 0
    assert end['counts'] == {'success': 2, 'failed': 0, 'skipped': 0}
    assert end['duration'] >= 0
    assert end['stopped_by_user'] is False
    assert end['login_failed'] is False


def test_one_failure_records_step_and_traceback(main_module, run, tmp_path):
    failed = main_module.CxrResult(False, '選取 Absent', error_type='TimeoutException',
                                   error_message='boom', traceback='Traceback (most recent call last): ...')
    run(FakeExcel(['HAP 1', 'HAP 2']), FakeAutomator(main_module, {'HAP 2': failed}))
    records = read_run_log(tmp_path)
    case = records[2]
    assert case['outcome'] == 'failed'
    assert case['step'] == '選取 Absent'
    assert case['error_type'] == 'TimeoutException'
    assert case['error_message'] == 'boom'
    assert case['traceback'].startswith('Traceback')
    assert records[-1]['counts'] == {'success': 1, 'failed': 1, 'skipped': 0}


def test_unknown_country_is_skipped(main_module, run, tmp_path):
    run(FakeExcel(['XYZ 4']), FakeAutomator(main_module))
    case = read_run_log(tmp_path)[1]
    assert case['outcome'] == 'skipped'
    assert case['skip_reason'] == 'unknown Country'
    assert case['country'] == '未知國家'


def test_missing_502_is_skipped(main_module, run, tmp_path):
    skipped = main_module.CxrResult(False, '選取 All 並進入 Manage Case',
                                    skip_reason='missing 502 Chest X-Ray Examination')
    run(FakeExcel(['HAP 1']), FakeAutomator(main_module, {'HAP 1': skipped}))
    records = read_run_log(tmp_path)
    assert records[1]['outcome'] == 'skipped'
    assert records[1]['skip_reason'] == 'missing 502 Chest X-Ray Examination'
    assert records[-1]['counts'] == {'success': 0, 'failed': 0, 'skipped': 1}


def test_login_failure_is_batch_level(main_module, run, tmp_path):
    run(FakeExcel(['HAP 1']), FakeAutomator(main_module, login_ok=False))
    records = read_run_log(tmp_path)
    assert [r['kind'] for r in records] == ['run_start', 'run_end']
    assert records[-1]['login_failed'] is True
    assert records[-1]['counts'] == {'success': 0, 'failed': 0, 'skipped': 0}


def test_user_stop_mid_batch(main_module, run, tmp_path):
    automator = FakeAutomator(main_module, on_run=lambda emed_no: main_module.stop_event.set())
    run(FakeExcel(['HAP 1', 'HAP 2', 'HAP 3']), automator)
    records = read_run_log(tmp_path)
    assert [r['kind'] for r in records] == ['run_start', 'case', 'run_end']
    assert records[-1]['stopped_by_user'] is True
    assert records[-1]['counts'] == {'success': 1, 'failed': 0, 'skipped': 0}


def test_records_before_a_crash_survive(main_module, run, tmp_path):
    def crash(emed_no):
        if emed_no == 'HAP 2':
            raise RuntimeError('crash')

    with pytest.raises(RuntimeError):
        run(FakeExcel(['HAP 1', 'HAP 2']), FakeAutomator(main_module, on_run=crash))
    records = read_run_log(tmp_path)
    assert [r['kind'] for r in records] == ['run_start', 'case']
    assert records[1]['emed_no'] == 'HAP 1'


def test_no_other_excel_cell_in_run_log(main_module, run, tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(['Patient Name', 'Date of Birth', 'eMedical No.', 'Passport'])
    ws.append(['Alice Secretname', '1990-01-02', 'HAP 47994319', 'PX998877'])
    ws.append(['Bob Hiddenname', '1985-03-04', 'CEAC 123', 'PX112233'])
    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(color='FF000000')  # the reader rejects the theme default font (xfail bug)
    wb.save(tmp_path / 'patients.xlsx')
    run(main_module.ExcelProcessor(), FakeAutomator(main_module), excel_name='patients.xlsx')
    text = next(Path(tmp_path / 'run_logs').glob('*.jsonl')).read_text(encoding='utf-8')
    assert 'HAP 47994319' in text and 'CEAC 123' in text
    for cell in ['Patient Name', 'Date of Birth', 'Passport', 'Alice Secretname', '1990-01-02',
                 'PX998877', 'Bob Hiddenname', '1985-03-04', 'PX112233']:
        assert cell not in text
