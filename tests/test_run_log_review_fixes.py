"""Run Log fixes from the spec #9 code review (ADR-0002: fakes only, no Live Portal)."""
import logging

from test_cxr_form_flow import automator, portal  # noqa: F401  (pytest fixtures)
from test_run_log import FakeAutomator, FakeExcel, read_run_log, run  # noqa: F401  (pytest fixtures)


def test_run_log_attribute_exists_before_a_run(main_module):
    manager = main_module.EmedicalWorkflowManager(FakeExcel([]), FakeAutomator(main_module))
    assert manager._run_log is None


def test_empty_excel_is_batch_level(main_module, run, tmp_path):  # noqa: F811
    run(FakeExcel([]), FakeAutomator(main_module))
    records = read_run_log(tmp_path)
    assert [r['kind'] for r in records] == ['run_start', 'run_end']
    assert records[0]['case_count'] == 0


def test_missing_excel_is_batch_level(main_module, tmp_path):
    manager = main_module.EmedicalWorkflowManager(FakeExcel(['HAP 1']), FakeAutomator(main_module))
    manager.set_gui_callbacks(*[lambda *a, **k: None] * 6)
    manager.start_workflow(str(tmp_path / 'missing.xlsx'), close_browser=False)
    records = read_run_log(tmp_path)
    assert [r['kind'] for r in records] == ['run_start', 'run_end']
    assert records[0]['case_count'] == 0
    assert records[0]['excel_file'] == 'missing.xlsx'


def test_non_selenium_error_becomes_failed_result(portal, automator, monkeypatch, main_module):  # noqa: F811
    def boom(*a, **k):
        raise ValueError('bad value')

    monkeypatch.setattr(main_module, 'write', boom)
    result = automator.run_cxr_exam('HAP 47994319', '澳大利亞')
    assert result.success is False
    assert result.skip_reason is None
    assert result.step
    assert result.error_type == 'ValueError'
    assert result.error_message == 'bad value'
    assert 'Traceback' in result.traceback and 'ValueError' in result.traceback


def test_stop_during_last_case_is_recorded(main_module, run, tmp_path):  # noqa: F811
    automator = FakeAutomator(main_module, on_run=lambda emed_no: main_module.stop_event.set())  # noqa: F811
    run(FakeExcel(['HAP 1']), automator)
    records = read_run_log(tmp_path)
    assert [r['kind'] for r in records] == ['run_start', 'case', 'run_end']
    assert records[-1]['stopped_by_user'] is True


def test_run_log_io_error_does_not_abort_batch(main_module, tmp_path, monkeypatch, caplog):
    monkeypatch.setattr(main_module, 'stop_event', __import__('threading').Event())
    (tmp_path / 'run_logs').write_text('a file, not a folder')
    successes = []
    manager = main_module.EmedicalWorkflowManager(FakeExcel(['HAP 1', 'HAP 2']), FakeAutomator(main_module))
    manager.set_gui_callbacks(lambda *a: None, lambda *a, **k: None, successes.append, lambda *a: None,
                              lambda: None, lambda: None)
    excel = tmp_path / 'cases.xlsx'
    excel.write_bytes(b'')
    with caplog.at_level(logging.ERROR):
        manager.start_workflow(str(excel), close_browser=False)
    assert successes == ['HAP 1', 'HAP 2']
    assert 'Run Log' in caplog.text
