"""Shared open-case steps (#14), with the fake helium from the form flow tests (ADR-0002: no Live Portal)."""
from test_cxr_form_flow import automator, portal  # noqa: F401  (pytest fixtures)


def test_open_case_searches_and_opens_manage_case(portal, automator):  # noqa: F811
    steps = []
    automator.open_case('HAP 47994319', steps.append)
    assert portal.writes == [('HAP 47994319', 'ID')]
    assert portal.clicks == ['Using Health Case Identifier', 'Search', 'All', 'Manage Case']
    assert steps == ['選擇 Health Case Identifier', '輸入 ID 並搜尋', '選取 All 並進入 Manage Case']


def test_form_flow_failure_in_open_case_keeps_step_name(portal, automator):  # noqa: F811
    portal.fail_on_click.add('Search')
    result = automator.run_cxr_exam('HAP 47994319', '澳大利亞')
    assert result.success is False
    assert result.step == '輸入 ID 並搜尋'
    assert portal.clicks[-1] == 'Close'
