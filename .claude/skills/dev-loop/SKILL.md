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
2. **Stop check** (hook for #8): before each fix round, check the stop conditions in [Stop conditions](#stop-conditions). When one fires, stop → Handoff.
3. `r = r + 1`. Call `debugger` with round `r` and `.loop/test-report-<r-1>.json`. Go to 1.

## 6. Review

1. `git add -A` as its own command (`.loop/` stays out: it is git-ignored).
2. Call `reviewer` with the issue number. It writes `.loop/review.json` with the tree of the staged change.
3. Read `.loop/review.json`:
   - `APPROVE`: go to step 7.
   - `REQUEST_CHANGES`: call `developer` with the summary as the changes to make, then go back to step 5 with the next round number. The review counts as a fix round for the stop conditions.

## 7. Commit

Run `git commit -m "<message>"` as its own command: plain `-m`, never `-a`, never chained after `git add`. The commit gate hook reruns the Tester and checks that `.loop/review.json` is `APPROVE` for the staged tree.

End the message with `Closes #<n>` and the attribution lines from the session's system reminder.

When the hook blocks the commit, read its reason. A failed Quality Gate goes back to step 5 with the next round; a tree mismatch goes back to step 6.

## 8. Pull request

1. `git push -u origin loop/<n>-<slug>`.
2. Load the `mattpocock-skills:pr` skill with the Skill tool and write the PR body from the issue, `.loop/plan.md` and the last Test Report. Include `Closes #<n>`.
3. `gh pr create --base main --head loop/<n>-<slug> --title "<issue title>" --body-file <file>`.

Done when the PR exists. Reply with the PR URL and the round count.

## Stop conditions

Placeholder, filled by #8. Fire on any of:

- 3 fix rounds done (`r == 3`) and the Tester still fails.
- The same set of failing gate locations in 2 rounds in a row.
- A change that needs the Live Portal.

## Handoff

Placeholder, filled by #8. For now: write `.loop/handoff.md` with the stop reason, the failed gates of the last Test Report and the content of `.loop/attempts.md`; report it to the user and end. Leave the branch as it is.
