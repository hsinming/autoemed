"""Tester hardening: file names with spaces, CRLF-only rewrites, deleted test files."""
import json
import subprocess
import sys
from pathlib import Path

import pytest

TESTER = Path(__file__).resolve().parents[1] / "scripts" / "quality_gate.py"


@pytest.fixture
def repo(tmp_path):
    """A temp git repo (no autocrlf) with a committed tests/test_kept.py."""
    subprocess.run(["git", "init", "-q"], cwd=tmp_path, check=True)
    subprocess.run(["git", "config", "core.autocrlf", "false"], cwd=tmp_path, check=True)
    (tmp_path / "tests").mkdir()
    (tmp_path / "tests" / "test_kept.py").write_bytes(b"def test_kept():\n    assert 1 == 1\n")
    subprocess.run(["git", "add", "-A"], cwd=tmp_path, check=True)
    subprocess.run(["git", "-c", "user.name=t", "-c", "user.email=t@t", "commit", "-q", "-m", "init"], cwd=tmp_path, check=True)
    return tmp_path


def run_tester(repo):
    subprocess.run(["git", "add", "-A"], cwd=repo, check=True)
    proc = subprocess.run([sys.executable, str(TESTER)], cwd=repo, capture_output=True, text=True)
    report = json.loads((repo / ".loop" / "test-report.json").read_text(encoding="utf-8"))
    return proc.returncode, {g["name"]: g for g in report["gates"]}


def test_syntax_error_in_file_name_with_space_is_found(repo):
    (repo / "bad name.py").write_text("x = 1\ndef f(:\n")
    code, gates = run_tester(repo)
    assert code != 0
    assert [f["location"] for f in gates["syntax"]["failures"]] == ["bad name.py:2"]


def test_crlf_only_rewrite_passes_protection(repo):
    (repo / "tests" / "test_kept.py").write_bytes(b"def test_kept():\r\n    assert 1 == 1\r\n")
    _code, gates = run_tester(repo)
    assert gates["protection"]["passed"] is True, gates["protection"]["failures"]


def test_deleted_test_file_fails_protection(repo):
    (repo / "tests" / "test_kept.py").unlink()
    code, gates = run_tester(repo)
    assert code != 0
    assert [f["location"] for f in gates["protection"]["failures"]] == ["tests/test_kept.py"]


def test_deleted_test_line_fails_protection(repo):
    (repo / "tests" / "test_kept.py").write_bytes(b"def test_kept():\n")
    code, gates = run_tester(repo)
    assert code != 0
    assert [f["location"] for f in gates["protection"]["failures"]] == ["tests/test_kept.py"]
