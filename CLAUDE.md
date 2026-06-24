# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

`autoemed` is a desktop automation tool that batch-processes Australian immigration health (eMedical) Chest X-Ray (CXR) normal case entries. It reads case IDs from an Excel file and drives Chrome via Selenium/Helium to fill out the eMedical web portal forms.

## Environment Setup

```bash
uv sync          # install/sync dependencies (preferred)
python main.py   # launch the Tkinter GUI
```

No test suite and no linter are configured in this project.

## Building a Standalone Executable

Use the PowerShell script (Windows target; dev environment is macOS):

```powershell
.\build_app.ps1
```

All Nuitka flags (`--standalone`, `--mingw64`, icon, output dir, plugins) are embedded as `nuitka-project:` comments at the top of `main.py` and are read automatically by Nuitka. Do not duplicate them in `build_app.ps1`.

## Architecture

Everything lives in a single file: `main.py`. Four classes wire together sequentially:

- **`ExcelProcessor`** — reads `.xlsx`, finds `"eMedical No."` column, filters out highlighted/colored rows (immigration clinic color-codes rows to skip), returns a plain list of case ID strings.

- **`EmedicalWebAutomator`** — wraps Helium/Selenium. `login()` opens Chrome, navigates to `EMEDICAL_URL`, and waits up to 300 seconds for the user to complete manual login and email-based 2FA until "Case search" appears. `automate_cxr_exam(emed_no, country)` navigates the portal form. **Country determines form flow**: Australia/NZ/Canada use a "Detailed radiology findings" multi-step path; USA (`CEAC` prefix) uses a simpler "Findings" path. Each step updates a `step` variable so failures log exactly where they occurred.

- **`EmedicalWorkflowManager`** — orchestrates the full batch: reads Excel → populates GUI → opens browser for manual login → loops over case IDs → calls `automate_cxr_exam` → tracks success/failure. Before processing each ID, `_normalize_emed_no()` inserts a space between the letter prefix and digits if missing (e.g. `HAP47994319` → `HAP 47994319`), which is the format the eMedical portal accepts. Country is detected from the normalized ID prefix (`HAP`/`TRN` → Australia, `NZER`/`NZHR` → NZ, `IME`/`UMI`/`UCI` → Canada, `CEAC` → USA). Decoupled from GUI via `set_gui_callbacks(...)`. Respects `stop_event` (a `threading.Event`) to allow mid-batch cancellation.

- **`EmedicalGUI`** — Tkinter UI. Runs the workflow in a daemon thread. Exposes a close-browser checkbox, three listboxes (all IDs / successes / failures), and Start/Stop buttons. No login credentials are entered here — the user logs in manually in the browser window.

**Known quirk:** `stop_event` is a module-level global and is never reset, so Stop can only be used once per process without restarting.

## Key Constants

- `VERSION` — version string defined in `main.py`. Update when bumping.
- `EMEDICAL_URL` — the portal URL; changes require reviewing the Selenium navigation selectors in `automate_cxr_exam`.

## Code Style Notes

- Comments and log messages are a mix of English and Traditional Chinese (繁體中文) — match the existing style in the surrounding code.
- `helium` functions (`click`, `write`, `find_all`, etc.) are imported explicitly from `helium`; don't replace with raw Selenium unless helium cannot handle the interaction.
