"""Stop check: does the failing set repeat between two Test Reports? (/dev-loop stop condition)

Usage: uv run python scripts/stop_check.py PREVIOUS_REPORT CURRENT_REPORT
Exit code 1 (stop) when both reports fail on the same set of (gate, location), 0 (continue) otherwise.
Prints the failing set of the current report, one "gate: location" per line.
"""
import json
import sys
from pathlib import Path


def failing_set(report_path):
    report = json.loads(Path(report_path).read_text(encoding="utf-8"))
    # message 不比對：同一處失敗但訊息略有不同（例如數值）仍算卡住
    return {(g["name"], f["location"]) for g in report["gates"] for f in g["failures"]}


def main(argv):
    prev, cur = failing_set(argv[0]), failing_set(argv[1])
    for gate, location in sorted(cur):
        print(f"{gate}: {location}")
    return 1 if cur and cur == prev else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
