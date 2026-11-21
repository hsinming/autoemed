"""Commit gate hook tests: pipe Claude Code PreToolUse hook JSON on stdin (Seam 2)."""
import json
import subprocess
import sys
from pathlib import Path

import pytest

HOOK = Path(__file__).resolve().parents[1] / "scripts" / "commit_gate.py"
SETTINGS = Path(__file__).resolve().parents[1] / ".claude" / "settings.json"


@pytest.fixture
def git_repo(tmp_path):
    """A temp git repo with one clean module and one passing test, all staged."""
    subprocess.run(["git", "init", "-q"], cwd=tmp_path, check=True)
    (tmp_path / "app.py").write_text("def add(a, b):\n    return a + b\n")
    (tmp_path / "test_app.py").write_text("from app import add\n\n\ndef test_add():\n    assert add(1, 2) == 3\n")
    subprocess.run(["git", "add", "-A"], cwd=tmp_path, check=True)
    return tmp_path


def staged_tree(repo):
    return subprocess.run(["git", "write-tree"], cwd=repo, capture_output=True, text=True, check=True).stdout.strip()


def write_verdict(repo, verdict, tree=None):
    (repo / ".loop").mkdir(exist_ok=True)
    (repo / ".loop" / "review.json").write_text(
        json.dumps({"verdict": verdict, "tree": tree or staged_tree(repo), "summary": "looks fine"})
    )


def run_hook(repo, command):
    payload = {"hook_event_name": "PreToolUse", "tool_name": "Bash", "cwd": str(repo), "tool_input": {"command": command}}
    return subprocess.run([sys.executable, str(HOOK)], cwd=repo, input=json.dumps(payload), capture_output=True, text=True)


def test_non_commit_command_passes_through(git_repo):
    proc = run_hook(git_repo, "git status && ls")
    assert proc.returncode == 0


def test_passing_gates_and_matching_approve_allow_commit(git_repo):
    write_verdict(git_repo, "APPROVE")
    proc = run_hook(git_repo, 'git commit -m "add thing"')
    assert proc.returncode == 0, proc.stderr


def test_failing_gate_blocks(git_repo):
    (git_repo / "test_app.py").write_text("from app import add\n\n\ndef test_add():\n    assert add(1, 2) == 4\n")
    subprocess.run(["git", "add", "-A"], cwd=git_repo, check=True)
    write_verdict(git_repo, "APPROVE")
    proc = run_hook(git_repo, "git commit -m x")
    assert proc.returncode == 2
    assert "tests" in proc.stderr


def test_missing_verdict_blocks(git_repo):
    proc = run_hook(git_repo, "git commit -m x")
    assert proc.returncode == 2
    assert "review.json" in proc.stderr


def test_request_changes_blocks(git_repo):
    write_verdict(git_repo, "REQUEST_CHANGES")
    proc = run_hook(git_repo, "git commit -m x")
    assert proc.returncode == 2
    assert "REQUEST_CHANGES" in proc.stderr


def test_tree_mismatch_blocks(git_repo):
    write_verdict(git_repo, "APPROVE", tree="0" * 40)
    proc = run_hook(git_repo, "git -C . commit -m x")
    assert proc.returncode == 2
    assert "tree" in proc.stderr


def test_hook_is_registered_in_project_settings():
    settings = json.loads(SETTINGS.read_text(encoding="utf-8"))
    [entry] = [e for e in settings["hooks"]["PreToolUse"] if "Bash" in e["matcher"].split("|")]
    assert any("commit_gate.py" in h["command"] for h in entry["hooks"])
    assert settings["enabledPlugins"]  # existing plugins kept
