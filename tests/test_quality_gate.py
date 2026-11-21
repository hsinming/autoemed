"""Tester CLI tests: run scripts/quality_gate.py against a temporary git repo."""
import json
import subprocess
import sys
from pathlib import Path

import pytest

TESTER = Path(__file__).resolve().parents[1] / "scripts" / "quality_gate.py"


@pytest.fixture
def git_repo(tmp_path):
    """A temp git repo with one clean module and one passing test."""
    subprocess.run(["git", "init", "-q"], cwd=tmp_path, check=True)
    (tmp_path / "app.py").write_text("def add(a, b):\n    return a + b\n")
    (tmp_path / "test_app.py").write_text("from app import add\n\n\ndef test_add():\n    assert add(1, 2) == 3\n")
    return tmp_path


def run_tester(repo):
    subprocess.run(["git", "add", "-A"], cwd=repo, check=True)
    proc = subprocess.run([sys.executable, str(TESTER)], cwd=repo, capture_output=True, text=True)
    report = json.loads((repo / ".loop" / "test-report.json").read_text(encoding="utf-8"))
    return proc.returncode, report


def gate(report, name):
    return next(g for g in report["gates"] if g["name"] == name)


def test_clean_repo_passes(git_repo):
    code, report = run_tester(git_repo)
    assert code == 0
    assert report["passed"] is True
    assert [g["name"] for g in report["gates"]] == ["syntax", "lint", "tests"]
    assert all(g["passed"] and g["failures"] == [] for g in report["gates"])
    assert len(report["tree"]) == 40
    assert report["timestamp"] and report["duration"] >= 0
    assert all(g["duration"] >= 0 for g in report["gates"])
    tests = gate(report, "tests")["tests"]
    assert [t["name"] for t in tests] == ["test_app.test_add"]
    assert tests[0]["duration"] >= 0


def test_syntax_error_fails_with_location(git_repo):
    (git_repo / "broken.py").write_text("x = 1\ndef f(:\n")
    code, report = run_tester(git_repo)
    assert code != 0
    assert report["passed"] is False
    [failure] = gate(report, "syntax")["failures"]
    assert failure["location"] == "broken.py:2"
    assert failure["message"]
    # all gates still run after one fails
    assert [g["name"] for g in report["gates"]] == ["syntax", "lint", "tests"]


def test_lint_error_fails_with_rule(git_repo):
    (git_repo / "app.py").write_text("import os\n\n\ndef add(a, b):\n    return a + b\n")
    code, report = run_tester(git_repo)
    assert code != 0
    assert gate(report, "syntax")["passed"] is True
    [failure] = gate(report, "lint")["failures"]
    assert failure["location"] == "app.py:1"
    assert "F401" in failure["message"]


def test_failing_test_reports_name_and_traceback(git_repo):
    (git_repo / "test_app.py").write_text("from app import add\n\n\ndef test_add():\n    assert add(1, 2) == 4\n")
    code, report = run_tester(git_repo)
    assert code != 0
    assert gate(report, "lint")["passed"] is True
    tests_gate = gate(report, "tests")
    [failure] = tests_gate["failures"]
    assert failure["location"] == "test_app.test_add"
    assert "assert 3 == 4" in failure["traceback"]
    assert tests_gate["tests"][0]["outcome"] == "failed"
