"""Stop check CLI tests: compare two Test Reports' failing (gate, location) sets."""
import json
import subprocess
import sys
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "stop_check.py"


def report(path, *failing):
    """Write a Test Report whose failing gates hold the given (gate, location, message) triples."""
    gates = {}
    for gate, location, message in failing:
        gates.setdefault(gate, []).append({"location": location, "message": message, "traceback": ""})
    data = {"passed": not failing, "tree": "abc", "gates": [
        {"name": name, "passed": False, "duration": 0.1, "failures": f} for name, f in gates.items()
    ] + [{"name": "syntax", "passed": True, "duration": 0.1, "failures": []}]}
    path.write_text(json.dumps(data))
    return str(path)


def run(prev, cur):
    return subprocess.run([sys.executable, str(SCRIPT), prev, cur], capture_output=True, text=True)


def test_same_failing_set_stops(tmp_path):
    prev = report(tmp_path / "a.json", ("tests", "tests/test_x.py::test_a", "assert 1 == 2"))
    cur = report(tmp_path / "b.json", ("tests", "tests/test_x.py::test_a", "assert 1 == 3"))
    proc = run(prev, cur)
    assert proc.returncode == 1
    assert "tests: tests/test_x.py::test_a" in proc.stdout


def test_different_failing_set_continues(tmp_path):
    prev = report(tmp_path / "a.json", ("tests", "tests/test_x.py::test_a", "boom"))
    cur = report(tmp_path / "b.json", ("tests", "tests/test_x.py::test_a", "boom"), ("lint", "main.py:3", "F401"))
    assert run(prev, cur).returncode == 0


def test_same_location_in_another_gate_is_different(tmp_path):
    prev = report(tmp_path / "a.json", ("syntax", "main.py:3", "x"))
    cur = report(tmp_path / "b.json", ("lint", "main.py:3", "x"))
    assert run(prev, cur).returncode == 0


def test_two_passing_reports_continue(tmp_path):
    assert run(report(tmp_path / "a.json"), report(tmp_path / "b.json")).returncode == 0
