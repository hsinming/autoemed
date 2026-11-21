"""Read-only Smoke Test (#15), with the fake helium from the form flow tests (ADR-0002: no Live Portal)."""
from selenium.common.exceptions import TimeoutException

from test_cxr_form_flow import automator, portal  # noqa: F401  (pytest fixtures)

# The read-only guard: adding any action here needs human review (#10).
ALLOWED_CLICKS = ['Using Health Case Identifier', 'Search', 'All', 'Manage Case', 'Close']


def test_smoke_pass_makes_only_read_only_actions(portal, automator, monkeypatch):  # noqa: F811
    monkeypatch.setattr(automator, 'login', lambda: True)
    result = automator.run_smoke_test('HAP47994319')
    assert result.success is True
    assert portal.writes == [('HAP 47994319', 'ID')]
    assert portal.clicks == ALLOWED_CLICKS


def test_smoke_missing_502_exam_fails_and_closes(portal, automator, monkeypatch):  # noqa: F811
    monkeypatch.setattr(automator, 'login', lambda: True)
    portal.missing.add('502 Chest X-Ray Examination')
    result = automator.run_smoke_test('HAP 47994319')
    assert result.success is False
    assert result.step == '檢查 502 Chest X-Ray Examination'
    assert portal.clicks == ALLOWED_CLICKS


def test_smoke_selector_timeout_fails_at_step_and_still_closes(portal, automator, monkeypatch, main_module):  # noqa: F811
    monkeypatch.setattr(automator, 'login', lambda: True)

    def wait_until(condition, *a, **k):
        raise TimeoutException('Select:')

    monkeypatch.setattr(main_module, 'wait_until', wait_until)
    result = automator.run_smoke_test('HAP 47994319')
    assert result.success is False
    assert result.step == '輸入 ID 並搜尋'
    assert result.error_type == 'TimeoutException'
    assert portal.clicks == ['Using Health Case Identifier', 'Search', 'Close']


def test_smoke_rejects_empty_emed_no_before_browser(portal, automator, monkeypatch):  # noqa: F811
    calls = []
    monkeypatch.setattr(automator, 'login', lambda: calls.append('login'))
    result = automator.run_smoke_test('   ')
    assert result.success is False
    assert calls == [] and portal.clicks == [] and portal.writes == []


def test_smoke_login_failure_fails_without_clicks(portal, automator, monkeypatch):  # noqa: F811
    monkeypatch.setattr(automator, 'login', lambda: False)
    result = automator.run_smoke_test('HAP 47994319')
    assert result.success is False
    assert result.step == '登入'
    assert portal.clicks == []


def test_smoke_login_exception_fails_at_login_step(portal, automator, monkeypatch):  # noqa: F811
    def login():
        raise RuntimeError('chromedriver missing')

    monkeypatch.setattr(automator, 'login', login)
    result = automator.run_smoke_test('HAP 47994319')
    assert result.success is False
    assert result.step == '登入'
    assert result.error_type == 'RuntimeError'


def test_smoke_pass_accepts_no_alert(portal, automator, monkeypatch, main_module):  # noqa: F811
    accepts = []

    def alert(*a, **k):
        a = portal.element('Alert')
        a.accept = lambda: accepts.append('accept')
        return a

    monkeypatch.setattr(main_module, 'Alert', alert)
    monkeypatch.setattr(automator, 'login', lambda: True)
    assert automator.run_smoke_test('HAP 47994319').success is True
    assert accepts == []
