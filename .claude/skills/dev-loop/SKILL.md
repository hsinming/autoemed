---
name: dev-loop
description: Run the Closed Loop on one Ready Issue, from plan to a pull request into main. Usage /dev-loop <issue number>.
disable-model-invocation: true
---

Drive the Closed Loop for the Ready Issue given as the argument. You run in the main session and call the role agents (`planner`, `developer`, `debugger`, `reviewer`) one level deep with the Agent tool; each agent gets the issue number and the paths it needs. The Tester (`scripts/quality_gate.py`) is the only proof that code passes (ADR-0001); nothing in the loop touches the Live Portal (ADR-0002).

Every artefact of the run goes in `.loop/` (git-ignored): `plan.md`, `test-report-<round>.json`, `attempts.md`, `review.json`.

**Stop → Handoff** means: go to [Handoff](#handoff) and end the loop there.

## 1. Gate the issue

Run `gh issue view <n> --json number,title,labels,state`. When the issue is closed or its labels lack `ready-for-agent`, refuse: reply "Issue #<n> is not a Ready Issue (needs the `ready-for-agent` label)" and end. Make no branch and no change.

## 2. Branch

1. `git status --porcelain` must be empty; when it is not, refuse and end.
2. Make the slug: the issue title, lower case, non-alphanumerics to `-`, at most 40 characters.
3. `git fetch origin main`, then `git switch -c loop/<n>-<slug> origin/main`.
4. Empty `.loop/` (create it if missing) so the run starts with no artefacts from an earlier run.

## 3. Plan

Call `planner` with the issue number. Read the verdict line in `.loop/plan.md`:

- `READY`: continue.
- `LIVE PORTAL: ...` or `BLOCKED: ...`: stop → Handoff with that line as the reason.

## 4. Develop

Call `developer` with the issue number.

## 5. Test and fix rounds

Set round `r = 0`.

1. Run the Tester: `uv run python scripts/quality_gate.py --report .loop/test-report-<r>.json`. Exit 0 means every Quality Gate passed: go to step 6.
2. **Stop check**: run the checks in [Stop conditions](#stop-conditions) on `.loop/test-report-<r>.json`. When one fires, stop → Handoff.
3. `r = r + 1`. Call `debugger` with round `r` and `.loop/test-report-<r-1>.json`. When its round in `.loop/attempts.md` has `Root cause: LIVE PORTAL: ...`, stop → Handoff with that line. Otherwise go to 1.

## 6. Review

1. `git add -A` as its own command (`.loop/` stays out: it is git-ignored).
2. Call `reviewer` with the issue number. It reads the staged change with `git diff --cached` (against HEAD) and writes `.loop/review.json` with the tree of the staged change.
3. Read `.loop/review.json`:
   - `APPROVE`: go to step 7.
   - `REQUEST_CHANGES`: call `developer` with the summary as the changes to make, then go back to step 5 with the next round number. The review counts as a fix round for the stop conditions.

## 7. Commit

Run `git commit -m "<message>"` as its own command: plain `-m`, never `-a`, never chained after `git add`. The commit gate hook refuses unstaged changes to tracked files, reruns the Tester and checks that `.loop/review.json` is `APPROVE` for the staged tree.

End the message with `Closes #<n>` and the attribution lines from the session's system reminder.

When the hook blocks the commit, read its reason. A failed Quality Gate goes back to step 5 with the next round; a tree mismatch goes back to step 6.

## 8. Pull request

1. `git push -u origin loop/<n>-<slug>`.
2. Load the `mattpocock-skills:pr` skill with the Skill tool and write the PR body from the issue, `.loop/plan.md` and the last Test Report. Include `Closes #<n>`.
3. `gh pr create --base main --head loop/<n>-<slug> --title "<issue title>" --body-file <file>`.

Done when the PR exists. Reply with the PR URL and the round count.

## Stop conditions

Check in this order after a failed Tester run in round `r`; the first that fires is the stop reason:

1. **Round limit**: `r == 3`. Reason: `3 fix rounds done and Quality Gates still fail`.
2. **Stuck**: when `r >= 1`, run `uv run python scripts/stop_check.py .loop/test-report-<r-1>.json .loop/test-report-<r>.json`. Exit 1 means both rounds fail on the same set of (gate, location). Reason: `same failing set in rounds <r-1> and <r>`. The script decides; take its exit code as the answer.
3. **Live Portal**: the plan verdict (step 3) is `LIVE PORTAL: ...`, or a Debugger round has `Root cause: LIVE PORTAL: ...` (step 5.3): the change needs `EMEDICAL_URL`, a selector or the portal flow (ADR-0002). Reason: that `LIVE PORTAL: ...` line.

Planner `BLOCKED: ...` also stops (step 3), with that line as the reason.

## Handoff

1. Write `.loop/handoff.md`:

   ```
   # Handoff: #<n> <issue title>

   **Stop reason:** <reason>
   **Branch:** loop/<n>-<slug> (local, not pushed) · **Fix rounds:** <r>

   ## Failed gates
   - <gate>: <location> — <message>   (one line per failure in the last Test Report)

   ## Last Test Report
   .loop/test-report-<r>.json · passed: <passed> · tree: <tree>
   - <gate>: pass|FAIL (<n> failures)   (one line per gate)

   ## Attempts per round
   <content of .loop/attempts.md, or "None.">

   ## Suggested next step
   <one or two lines: what a human checks or changes first>
   ```

   When the loop stopped at the plan, write "No Test Report: stopped at plan." under both gate sections and quote the plan's open questions as the next step.

   Only when the stop reason contains `LIVE PORTAL:` (for example the Debugger's `Root cause: LIVE PORTAL: ...`), append this section. Other stop reasons leave it out. You never run the Smoke Test yourself (ADR-0002); the human runs it.

   ```
   ## Smoke Test (human)
   This change touches `EMEDICAL_URL`, a selector or the portal flow. After the fix, run the Smoke Test on the Live Portal:

       uv run python main.py --smoke "<eMed No.>"

   Use an eMed No. that is safe to open. The Smoke Test is read-only, but it opens a real health case.
   ```
2. `gh issue comment <n> --body-file .loop/handoff.md`.
3. `gh issue edit <n> --remove-label ready-for-agent --add-label ready-for-human`. Step 1 then refuses the issue until a human puts `ready-for-agent` back.
4. Leave the branch as it is. Reply with the stop reason and the issue URL, and end.
