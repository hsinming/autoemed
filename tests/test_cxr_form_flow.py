"""CXR form flow per Country, with helium replaced by fakes (ADR-0002: no Live Portal)."""
import logging

import pytest
from selenium.common.exceptions import NoSuchElementException

DECLARATION = 'I declare that the chest X-ray examination report is a true and correct record of my findings.'
GRADE_A = ('A - No evidence of active TB, or changes consistent with old or inactive TB, '
           'or changes suggestive of other significant diseases identified.')

START = [
    'Using Health Case Identifier', 'Search', 'All', 'Manage Case', '502 Chest X-Ray Examination',
]
DETAILED = ['Detailed radiology findings', 'Normal', 'Normal', 'Absent', 'No']
FINISH_GRADING = ['Next', 'Next', 'Prepare for grading', DECLARATION, GRADE_A, 'Submit Exam', 'Close']


class FakePortal:
    """Records clicks and writes; nothing is selected, and every element exists unless listed in `missing`."""

    def __init__(self, missing=(), fail_on_click=()):
        self.missing = set(missing)
        self.fail_on_click = set(fail_on_click)
        self.clicks = []
        self.writes = []

    def element(self, label='', *args, **kwargs):
        portal = self

        class FakeElement:
            def exists(self):
                return label not in portal.missing

            def is_selected(self):
                return False

            def is_checked(self):
                return False

            def is_enabled(self):
                return True

            def accept(self):
                pass

        e = FakeElement()
        e.label = label
        return e

    def click(self, element):
        if element.label in self.fail_on_click:
            raise NoSuchElementException(element.label)
        self.clicks.append(element.label)

    def write(self, text, into=None):
        self.writes.append((text, into.label))


@pytest.fixture
def portal(main_module, monkeypatch):
    fake = FakePortal()
    for name in ('Text', 'Button', 'RadioButton', 'CheckBox', 'TextField', 'Alert'):
        monkeypatch.setattr(main_module, name, fake.element)
    monkeypatch.setattr(main_module, 'click', fake.click)
    monkeypatch.setattr(main_module, 'write', fake.write)
    monkeypatch.setattr(main_module, 'wait_until', lambda *a, **k: None)
    monkeypatch.setattr(main_module, 'find_all', lambda e: [fake.element('Normal'), fake.element('Normal')])
    monkeypatch.setattr(main_module, 'sleep', lambda s: None)
    monkeypatch.setattr(main_module, 'start_chrome', lambda *a, **k: pytest.fail('real browser started'))
    return fake


@pytest.fixture
def automator(main_module):
    return main_module.EmedicalWebAutomator(main_module.EMEDICAL_URL)


@pytest.mark.parametrize('country', ['澳大利亞', '紐西蘭'])
def test_detailed_radiology_flow(portal, automator, country):
    assert automator.automate_cxr_exam('HAP 47994319', country) is True
    assert portal.writes == [('HAP 47994319', 'ID')]
    assert portal.clicks == START + DETAILED + FINISH_GRADING


def test_canada_flow_adds_special_findings(portal, automator):
    assert automator.automate_cxr_exam('IME 12345678', '加拿大') is True
    assert portal.clicks == (START + DETAILED + ['Next', 'None of the following are present']
                             + FINISH_GRADING)


def test_usa_findings_flow(portal, automator):
    assert automator.automate_cxr_exam('CEAC 2024123456', '美國') is True
    assert portal.clicks == START + [
        'Findings', 'Normal', 'Next', 'Next', 'Prepare for declaration', DECLARATION, 'Submit Exam', 'Close',
    ]


def test_missing_502_exam_fails_and_closes_case(portal, automator):
    portal.missing.add('502 Chest X-Ray Examination')
    assert automator.automate_cxr_exam('HAP 47994319', '澳大利亞') is False
    assert portal.clicks == START[:-1] + ['Close']


def test_failure_logs_current_step(portal, automator, caplog):
    portal.fail_on_click.add('Absent')
    with caplog.at_level(logging.ERROR):
        assert automator.automate_cxr_exam('HAP 47994319', '澳大利亞') is False
    assert 'step: 選取 Absent' in caplog.text
    assert portal.clicks[-1] == 'Close'
