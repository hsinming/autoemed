"""Commit gate: Claude Code PreToolUse hook on Bash and PowerShell (ADR-0001).

Reads the hook JSON on stdin. For a `git commit` command it refuses unstaged changes to tracked files,
reruns the Tester live at the repo root, then checks `.loop/review.json` ({"verdict", "tree", "summary"}):
the verdict must be APPROVE for the staged tree.
Exit 0 allows the command; exit 2 blocks it with the reason on stderr. Any error blocks (fail closed).
"""
import json
import shlex
import subprocess
import sys
from pathlib import Path

TESTER = Path(__file__).resolve().parent / "quality_gate.py"
VERDICT = Path(".loop") / "review.json"
GIT_OPTS_WITH_VALUE = {"-C", "-c", "--git-dir", "--work-tree", "--namespace", "--exec-path", "--config-env"}


def is_commit(command):
    """True when a segment of command (split on && ; | & and newlines) runs `git [global options] commit`."""
    # shortcut: wrappers such as `VAR=1 git commit`, `bash -c "git commit"` or git aliases are not seen;
    # upgrade to a git pre-commit hook if agents commit that way.
    lex = shlex.shlex(command.replace("\n", ";"), posix=True, punctuation_chars=True)
    lex.whitespace_split = True
    lex.escape = ""  # keep Windows backslash paths intact
    try:
        tokens = list(lex)
    except ValueError:  # unbalanced quotes: let check() decide rather than skip the gate
        return "commit" in command
    segments, seg = [], []
    for tok in tokens + [";"]:
        if tok and all(c in "();<>|&" for c in tok):
            segments.append(seg)
            seg = []
        else:
            seg.append(tok)
    for seg in segments:
        if not seg or Path(seg[0].replace("\\", "/")).name.lower() not in ("git", "git.exe"):
            continue
        args = iter(seg[1:])
        for arg in args:
            if arg in GIT_OPTS_WITH_VALUE:
                next(args, None)
            elif not arg.startswith("-"):
                if arg == "commit":
                    return True
                break
    return False


def check(cwd):
    """Return the reasons to block a commit in the repo that contains cwd (empty list = allow)."""
    cwd = subprocess.run(["git", "rev-parse", "--show-toplevel"], cwd=cwd, capture_output=True, text=True, check=True).stdout.strip()
    # shortcut: untracked files are tested by the Tester but not committed unless staged; upgrade to testing a
    # checkout of the staged tree if that gap bites.
    if subprocess.run(["git", "diff", "--quiet"], cwd=cwd).returncode:
        return ["Tracked files have unstaged changes, so the tested state differs from the staged one; "
                "`git add` them (or revert them) and review again."]
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
    try:
        payload = json.load(sys.stdin)
        if not is_commit(payload.get("tool_input", {}).get("command", "")):
            return 0
        reasons = check(payload.get("cwd") or ".")
    except Exception as e:  # fail closed: exit 1 would let the command run
        reasons = [f"Commit gate error: {e!r}"]
    if reasons:
        print("Commit blocked by commit gate:\n- " + "\n- ".join(reasons), file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
