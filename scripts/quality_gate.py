"""Tester: run every Quality Gate on the cwd git repo and write one Test Report (ADR-0001).

Usage: uv run python scripts/quality_gate.py [--report PATH]
Exit code 0 when every gate passes, 1 otherwise. Every gate runs even after one fails.
"""
import argparse
import json
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path

DEFAULT_REPORT = Path(".loop") / "test-report.json"
LINT_RULES = "E4,E7,E9,F"  # ruff 的 default rule set，固定寫死，避免新版 ruff 預設規則變動


def git(*args):
    return subprocess.run(["git", *args], capture_output=True, text=True, check=True).stdout


def failure(location, message, traceback=""):
    return {"location": location, "message": message, "traceback": traceback}


def gate_syntax():
    """Compile every tracked (or new, not ignored) .py file."""
    files = git("ls-files", "--cached", "--others", "--exclude-standard", "--", "*.py").split()
    failures = []
    for name in files:
        path = Path(name)
        if not path.exists():  # deleted in the working tree
            continue
        try:
            compile(path.read_bytes(), name, "exec")
        except SyntaxError as e:
            failures.append(failure(f"{name}:{e.lineno}", e.msg))
    return {"failures": failures}


def gate_lint():
    proc = subprocess.run(
        [sys.executable, "-m", "ruff", "check", "--select", LINT_RULES, "--output-format", "json", "."],
        capture_output=True, text=True, encoding="utf-8",
    )
    try:
        issues = json.loads(proc.stdout)
    except json.JSONDecodeError:
        return {"failures": [failure("ruff", "ruff did not run", proc.stderr)]}
    cwd = Path.cwd()
    return {"failures": [
        failure(f"{Path(i['filename']).relative_to(cwd).as_posix()}:{i['location']['row']}", f"{i['code']} {i['message']}")
        for i in issues
    ]}


def gate_tests():
    """Run pytest and read per-test results from its JUnit XML."""
    with tempfile.TemporaryDirectory() as tmp:
        xml_path = Path(tmp) / "junit.xml"
        proc = subprocess.run(
            [sys.executable, "-m", "pytest", "-q", "-p", "no:cacheprovider", f"--junitxml={xml_path}"],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        cases = ET.parse(xml_path).getroot().iter("testcase") if xml_path.exists() else []
        tests, failures = [], []
        for case in cases:
            name = f"{case.get('classname')}.{case.get('name')}"
            outcome = "passed"
            for tag in ("failure", "error", "skipped"):
                el = case.find(tag)
                if el is not None:
                    outcome = {"failure": "failed", "error": "error", "skipped": "skipped"}[tag]
                    if tag != "skipped":
                        failures.append(failure(name, el.get("message", ""), el.text or ""))
                    break
            tests.append({"name": name, "outcome": outcome, "duration": float(case.get("time", 0))})
    # 0 = all passed, 5 = no tests collected; anything else without a parsed failure is a pytest crash
    if proc.returncode not in (0, 5) and not failures:
        failures.append(failure("pytest", f"pytest exit code {proc.returncode}", proc.stdout[-4000:] + proc.stderr[-4000:]))
    return {"failures": failures, "tests": tests}


def gate_protection():
    """Fail when the diff of tests/ against HEAD removes or changes an existing line (ADR-0001)."""
    if subprocess.run(["git", "rev-parse", "-q", "--verify", "HEAD"], capture_output=True).returncode:
        return {"failures": []}  # no commit yet: no existing test lines to protect
    failures = []
    for line in git("diff", "HEAD", "--numstat", "--no-renames", "--", "tests/").splitlines():
        _added, removed, name = line.split("\t", 2)
        if removed != "0":  # a changed line counts as removed + added; "-" means a binary change
            failures.append(failure(name, f"{removed} existing test line(s) removed or changed; only a human may do this"))
    return {"failures": failures}


# Gates run in this order. Add a new Quality Gate by appending (name, fn); fn returns {"failures": [...], ...extra}.
GATES = [
    ("syntax", gate_syntax),
    ("lint", gate_lint),
    ("tests", gate_tests),
    ("protection", gate_protection),
]


def run(report_path):
    start = time.perf_counter()
    gates = []
    for name, fn in GATES:
        t0 = time.perf_counter()
        result = fn()
        gates.append({"name": name, "passed": not result["failures"], "duration": round(time.perf_counter() - t0, 3), **result})
    report = {
        "passed": all(g["passed"] for g in gates),
        "tree": git("write-tree").strip(),  # tree hash of the staged state; the commit hook compares this
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "duration": round(time.perf_counter() - start, 3),
        "gates": gates,
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    return report


def main():
    parser = argparse.ArgumentParser(description="Run the Quality Gates and write a Test Report.")
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT, help=f"Test Report path (default: {DEFAULT_REPORT})")
    report = run(parser.parse_args().report)
    for g in report["gates"]:
        print(f"{'PASS' if g['passed'] else 'FAIL'} {g['name']} ({g['duration']}s)")
        for f in g["failures"]:
            print(f"  {f['location']}: {f['message']}")
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
