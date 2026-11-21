"""CXR form flow result value, with the fake helium from the form flow tests (ADR-0002: no Live Portal)."""
from test_cxr_form_flow import automator, portal  # noqa: F401  (pytest fixtures)


def test_success_returns_last_step(portal, automator):  # noqa: F811
    result = automator.run_cxr_exam('HAP 47994319', '澳大利亞')
    assert result.success is True
    assert result.step == 'Close'
    assert result.skip_reason is None
    assert result.error_type is None


def test_missing_502_exam_returns_skip_with_reason(portal, automator):  # noqa: F811
    portal.missing.add('502 Chest X-Ray Examination')
    result = automator.run_cxr_exam('HAP 47994319', '澳大利亞')
    assert result.success is False
    assert result.skip_reason == 'missing 502 Chest X-Ray Examination'
    assert result.error_type is None


def test_failure_returns_step_error_and_traceback(portal, automator):  # noqa: F811
    portal.fail_on_click.add('Absent')
    result = automator.run_cxr_exam('HAP 47994319', '澳大利亞')
    assert result.success is False
    assert result.step == '選取 Absent'
    assert result.skip_reason is None
    assert result.error_type == 'NoSuchElementException'
    assert 'Absent' in result.error_message
    assert 'Traceback' in result.traceback and 'NoSuchElementException' in result.traceback


def test_gui_lists_follow_result_success(main_module, tmp_path):
    results = {
        'HAP 1': main_module.CxrResult(True, 'Close'),
        'HAP 2': main_module.CxrResult(False, '選取 Absent', error_type='TimeoutException'),
        'HAP 3': main_module.CxrResult(False, '選取 All 並進入 Manage Case',
                                      skip_reason='missing 502 Chest X-Ray Examination'),
    }

    class FakeExcel:
        def extract_emedical_no(self, path):
            return ['HAP1', 'HAP 2', 'HAP 3', 'XYZ 4']

    class FakeAutomator:
        def login(self):
            return True

        def run_cxr_exam(self, emed_no, country):
            return results[emed_no]

    successes, failures = [], []
    manager = main_module.EmedicalWorkflowManager(FakeExcel(), FakeAutomator())
    manager.set_gui_callbacks(lambda *a: None, lambda *a, **k: None, successes.append, failures.append,
                              lambda: None, lambda: None)
    excel = tmp_path / 'cases.xlsx'
    excel.write_bytes(b'')
    manager.start_workflow(str(excel), close_browser=False)
    assert successes == ['HAP1']
    assert failures == ['HAP 2', 'HAP 3', 'XYZ 4']
