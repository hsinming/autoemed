# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

`autoemed` is a desktop automation tool that batch-processes Australian immigration health (eMedical) Chest X-Ray (CXR) normal case entries. It reads case IDs from an Excel file and drives Chrome via Selenium/Helium to fill out the eMedical web portal forms.

## Environment Setup

```bash
uv sync                    # install/sync dependencies
uv run python main.py      # launch the Tkinter GUI
uv run python scripts/quality_gate.py   # Tester: syntax, lint (ruff E4,E7,E9,F), pytest → .loop/test-report.json
```

**Smoke Test** (human only, Live Portal, read-only): `uv run python main.py --smoke "<eMed No.>"`. Log in and complete 2FA in Chrome; it searches the case, opens Manage Case, checks "502 Chest X-Ray Examination" and clicks Close. It prints one `PASS`/`FAIL | step: ...` line and exits 0/1. Never run it from an agent or `/dev-loop` (ADR-0002). `tests/test_smoke.py` guards the allowed actions; adding one needs human review.

Tests live in `tests/` and never touch the Live Portal (ADR-0002). Import `main` only through the `main_module` fixture, which sets a temp cwd because `main` writes `log.txt` on import.

## Closed Loop

- **Tester**: `uv run python scripts/quality_gate.py [--report PATH]` runs the Quality Gates (syntax, lint, tests, test protection) and writes the Test Report (default `.loop/test-report.json`). Exit 0 means all gates pass; it is the only proof that code passes (ADR-0001). Test protection fails on any changed or removed line under `tests/` versus `HEAD`: add tests, never edit existing ones.
- **Commit gate**: a `PreToolUse` hook (`scripts/commit_gate.py`) reruns the Tester on every `git commit` from Claude Code and needs `.loop/review.json` = `{"verdict": "APPROVE", "tree": "<git write-tree of the staged change>", "summary": "..."}`. Commit as: `git add`, then the review verdict, then a separate plain `git commit -m`. It blocks when tracked files have unstaged changes, and blocks on any internal error (fail closed).
- **Strict xfail**: `tests/test_excel_reader.py` marks the theme-font bug `xfail(strict=True)`. When that bug is fixed the test XPASSes and fails the tests gate; a human must remove the marker, because test protection forbids agents to change existing test lines.
- **`/dev-loop <issue>`**: runs the Closed Loop on one Ready Issue (`ready-for-agent`): branch `loop/<issue>-<slug>` from `main`, then the `planner`, `developer`, `debugger` and `reviewer` agents in `.claude/agents/`, and ends in a PR into `main` or a Handoff. Run artefacts go in `.loop/` (git-ignored).

## Building a Standalone Executable

Use the PowerShell script (Windows target; dev environment is macOS):

```powershell
.\build_app.ps1
```

All Nuitka flags (`--standalone`, `--mingw64`, icon, output dir, plugins) are embedded as `nuitka-project:` comments at the top of `main.py` and are read automatically by Nuitka. Do not duplicate them in `build_app.ps1`.
## Architecture

Everything lives in a single file: `main.py`. Four classes wire together sequentially:

- **`ExcelProcessor`** — reads `.xlsx`, finds `"eMedical No."` column, filters out highlighted/colored rows (immigration clinic color-codes rows to skip), returns a plain list of case ID strings.

- **`EmedicalWebAutomator`** — wraps Helium/Selenium. `login()` opens Chrome, navigates to `EMEDICAL_URL`, and waits up to 300 seconds for the user to complete manual login and email-based 2FA until "Case search" appears. `run_cxr_exam(emed_no, country)` navigates the portal form and returns a `CxrResult` (success, last step, skip reason or error type, message and traceback); any exception becomes a failed result. `automate_cxr_exam` is only a bool wrapper kept for the #5 form-flow tests. `open_case(emed_no, set_step)` holds the search-and-Manage-Case steps shared by `run_cxr_exam` and `run_smoke_test`; `run_smoke_test(emed_no)` is the read-only Smoke Test behind `--smoke` (login, `open_case`, check 502, Close) and returns a `CxrResult`. **Country determines form flow**: Australia/NZ/Canada use a "Detailed radiology findings" multi-step path; USA (`CEAC` prefix) uses a simpler "Findings" path. Each step updates a `step` variable so failures log exactly where they occurred.

- **`EmedicalWorkflowManager`** — orchestrates the full batch: reads Excel → populates GUI → opens browser for manual login → loops over case IDs → calls `run_cxr_exam` → tracks success/failure and writes a Run Log (`run_logs/*.jsonl` next to `log.txt`). Before processing each ID, `_normalize_emed_no()` inserts a space between the letter prefix and digits if missing (e.g. `HAP47994319` → `HAP 47994319`), which is the format the eMedical portal accepts. Country is detected from the normalized ID prefix (`HAP`/`TRN` → Australia, `NZER`/`NZHR` → NZ, `IME`/`UMI`/`UCI` → Canada, `CEAC` → USA). Decoupled from GUI via `set_gui_callbacks(...)`. Respects `stop_event` (a `threading.Event`) to allow mid-batch cancellation.

- **`EmedicalGUI`** — Tkinter UI. Runs the workflow in a daemon thread. Exposes a close-browser checkbox, three listboxes (all IDs / successes / failures), and Start/Stop buttons. No login credentials are entered here — the user logs in manually in the browser window.

**Known quirk:** `stop_event` is a module-level global and is never reset, so Stop can only be used once per process without restarting.

## Key Constants

- `VERSION` — version string defined in `main.py`. Update when bumping.
- `EMEDICAL_URL` — the portal URL; changes require reviewing the Selenium navigation selectors in `run_cxr_exam`.

## Code Style Notes

- Comments and log messages are a mix of English and Traditional Chinese (繁體中文) — match the existing style in the surrounding code.
- `helium` functions (`click`, `write`, `find_all`, etc.) are imported explicitly from `helium`; don't replace with raw Selenium unless helium cannot handle the interaction.

## Agent skills

### Issue tracker

Issues live in GitHub Issues for `hsinming/autoemed`, via the `gh` CLI. See `docs/agents/issue-tracker.md`.

### Triage labels

The five default triage labels, plus priority (`P0`–`P2`) and source labels. See `docs/agents/triage-labels.md`.

### Run Log triage

Triage of a Run Log or a user report: grouping, priority, source labels, redaction. See `docs/agents/run-log-triage.md`.

### Domain docs

Single-context: `GLOSSARY.md` and `docs/adr/` at the repo root. See `docs/agents/domain.md`.
