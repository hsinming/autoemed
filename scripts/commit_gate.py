"""Commit gate: Claude Code PreToolUse hook on Bash (ADR-0001).

Reads the hook JSON on stdin. For a `git commit` command it reruns the Tester live, then checks
`.loop/review.json` ({"verdict", "tree", "summary"}): the verdict must be APPROVE for the staged tree.
Exit 0 allows the command; exit 2 blocks it with the reason on stderr.
"""
import json
import re
import subprocess
import sys
from pathlib import Path

TESTER = Path(__file__).resolve().parent / "quality_gate.py"
VERDICT = Path(".loop") / "review.json"
# `git [global options] commit`; global options such as -C <path> may sit between git and commit
COMMIT_RE = re.compile(r"\bgit\s+(?:-[cC]\s+\S+\s+|-\S+\s+)*commit\b")


def check(cwd):
    """Return the reasons to block a commit in cwd (empty list = allow)."""
    # shortcut: gates and tree see the index as it is now, so `git commit -a` or `git add ... && git commit`
    # are checked before their files are staged; upgrade to a git pre-commit hook if agents use those forms.
    tester = subprocess.run([sys.executable, str(TESTER)], cwd=cwd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    reasons = []
    if tester.returncode != 0:
        reasons.append("Quality Gates failed (see .loop/test-report.json):\n" + tester.stdout.strip())
    tree = subprocess.run(["git", "write-tree"], cwd=cwd, capture_output=True, text=True, check=True).stdout.strip()
    try:
        review = json.loads((Path(cwd) / VERDICT).read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        return reasons + [f"No readable review verdict at {VERDICT.as_posix()} ({e}); run the Reviewer first."]
    if review.get("verdict") != "APPROVE":
        reasons.append(f"Review verdict is {review.get('verdict')!r}, not 'APPROVE': {review.get('summary', '')}")
    if review.get("tree") != tree:
        reasons.append(f"Review tree {review.get('tree')} does not match staged tree {tree}; review the staged changes again.")
    return reasons


def main():
    payload = json.load(sys.stdin)
    if not COMMIT_RE.search(payload.get("tool_input", {}).get("command", "")):
        return 0
    reasons = check(payload.get("cwd") or ".")
    if reasons:
        print("Commit blocked by commit gate:\n- " + "\n- ".join(reasons), file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
