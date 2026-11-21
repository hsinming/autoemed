"""Commit gate hardening: fail closed, exact `git commit` matching, repo root, unstaged edits."""
import json
import subprocess
import sys
from pathlib import Path

import pytest

HOOK = Path(__file__).resolve().parents[1] / "scripts" / "commit_gate.py"


@pytest.fixture
def repo(tmp_path):
    """A temp git repo with one clean module and one passing test, all staged."""
    subprocess.run(["git", "init", "-q"], cwd=tmp_path, check=True)
    (tmp_path / "app.py").write_text("def add(a, b):\n    return a + b\n")
    (tmp_path / "test_app.py").write_text("from app import add\n\n\ndef test_add():\n    assert add(1, 2) == 3\n")
    subprocess.run(["git", "add", "-A"], cwd=tmp_path, check=True)
    return tmp_path


def approve(repo):
    tree = subprocess.run(["git", "write-tree"], cwd=repo, capture_output=True, text=True, check=True).stdout.strip()
    (repo / ".loop").mkdir(exist_ok=True)
    (repo / ".loop" / "review.json").write_text(json.dumps({"verdict": "APPROVE", "tree": tree, "summary": "ok"}))


def hook(cwd, command):
    payload = {"hook_event_name": "PreToolUse", "tool_name": "Bash", "cwd": str(cwd), "tool_input": {"command": command}}
    return subprocess.run([sys.executable, str(HOOK)], cwd=cwd, input=json.dumps(payload), capture_output=True, text=True)


def test_error_blocks_instead_of_failing_open(tmp_path):
    proc = hook(tmp_path, "git commit -m x")  # not a git repo: the gate cannot check
    assert proc.returncode == 2
    assert proc.stderr.strip()


def test_malformed_payload_blocks(tmp_path):
    proc = subprocess.run([sys.executable, str(HOOK)], cwd=tmp_path, input="not json", capture_output=True, text=True)
    assert proc.returncode == 2


@pytest.mark.parametrize("command", [
    "git commit-tree 4b825dc -m x",
    'echo "git commit"',
    'gh pr create --title t --body "this fixes the git commit gate"',
    "git log --grep commit",
    "git status && ls",
])
def test_non_commit_commands_pass_through(repo, command):
    assert hook(repo, command).returncode == 0  # no verdict exists, so a match would block


@pytest.mark.parametrize("command", [
    "git.exe commit -m x",
    '"C:/Program Files/Git/cmd/git.exe" commit -m x',
    r'"C:\Program Files\Git\cmd\git.exe" commit -m x',
    'git -C "some dir" commit -m x',
    "git -c user.name=x --no-pager commit -m x",
    "git status && git commit -m x",
    "git status; git commit -m x",
    "git status\ngit commit -m x",
    "& git commit -m x",
])
def test_commit_invocations_are_gated(repo, command):
    proc = hook(repo, command)  # no verdict: a gated commit is blocked
    assert proc.returncode == 2
    assert "review.json" in proc.stderr


def test_subdirectory_cwd_uses_repo_root(repo):
    approve(repo)
    sub = repo / "sub"
    sub.mkdir()
    proc = hook(sub, "git commit -m x")
    assert proc.returncode == 0, proc.stderr


def test_unstaged_tracked_changes_block(repo):
    approve(repo)
    (repo / "app.py").write_text("def add(a, b):\n    return a + b + 0\n")
    proc = hook(repo, "git commit -m x")
    assert proc.returncode == 2
    assert "unstaged" in proc.stderr
