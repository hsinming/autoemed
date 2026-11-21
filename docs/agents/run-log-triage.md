# Run Log triage

How `/triage` turns a Run Log, a user description, or both into GitHub issues. Labels are in `triage-labels.md`; `gh` commands are in `issue-tracker.md`.

## Run Log format

One JSON Lines file per batch run: `run_logs/run_YYYYMMDD_HHMMSS.jsonl`, beside `log.txt`. Each line has a `kind`:

- `run_start`: `version`, `excel_file`, `case_count`, `start_time`.
- `case`: `emed_no`, `country`, `outcome` (`success`, `failed`, `skipped`), `step`, `skip_reason`, `error_type`, `error_message`, `traceback`, `duration`.
- `run_end`: `counts`, `duration`, `stopped_by_user`, `login_failed`.

## Steps

1. **Group.** Take the `case` records with outcome `failed` or `skipped`. One pattern is one (`country`, `step`, `error_type` or `skip_reason`) tuple. Each pattern becomes one issue. A user description without a Run Log is one pattern.
2. **Prioritise.** `P0` when the batch fails as a whole: `login_failed` is true, or no case ran. `P1` for a case-failure pattern. `P2` for an improvement.
3. **Label the source.** `run-log` when the pattern comes from a Run Log, `user-report` when it comes from a user description, both when both.
4. **Search, then write.** Search open issues for the same Country, `step` and error type (`gh issue list --state open --search "<step> <error_type>"`). On a match, comment on that issue with the new counts. Only with no match, create an issue.
5. **Set the triage state.** A new issue gets `needs-triage`, or `needs-info` when the evidence cannot show the cause. `ready-for-agent` is a human decision, so triage leaves it off.

Done when every failure pattern in the input has exactly one new issue or one new comment.

## Public repo: redact

The repo is public. An issue body or comment holds only:

- counts (cases in the pattern, cases in the run),
- Country, `step`, error type or skip reason, app `version`,
- the `traceback` trimmed to the frames in `main.py` and the final exception line.

Copy no eMed No., in any form (with or without the space). `emed_no`, `excel_file` and `error_message` stay out, and so does any traceback line that contains an eMed No. Before each `gh` write, check the text against the eMed No. shape (regex `[A-Z]+ ?\d{6,}`) and remove every match.
