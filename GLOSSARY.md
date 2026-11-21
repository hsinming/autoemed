# autoemed

autoemed batch-enters normal Chest X-Ray results into the eMedical portal. This glossary also covers the agent-driven development loop that changes autoemed.

## eMedical domain

**eMed No.**:
The case identifier that the eMedical portal assigns to one health case, in the form `<letter prefix> <digits>` (e.g. `HAP 47994319`).
_Avoid_: case ID, emed number, HAP number

**Country**:
The destination country of a case, derived from the eMed No. prefix. It decides which form flow the portal shows.
_Avoid_: region, visa type

**Live Portal**:
The real eMedical website, with real health cases. A human must complete its 2FA login.
_Avoid_: production site, the website

## Development loop

**Closed Loop**:
The repeated cycle of plan, develop, test, debug, retest and review, which ends in a commit or a Handoff.
_Avoid_: pipeline, agent loop

**Tester**:
A deterministic script, not an agent, that runs the Quality Gates and writes the Test Report. Its result is the only proof that code passes.
_Avoid_: test agent, QA agent

**Quality Gate**:
One pass/fail check that the Tester runs before a commit: syntax, lint, tests, or test protection. The commit gate also checks the review verdict.
_Avoid_: check, CI step

**Test Report**:
The JSON file that the Tester writes after each run.
_Avoid_: test log, results

**Handoff**:
The stop of the Closed Loop when it needs a human, with a report of what failed and what the agents tried.
_Avoid_: escalation, abort

**Smoke Test**:
A read-only check against the Live Portal that a human starts and logs in to. It is never a Quality Gate.
_Avoid_: E2E test, live test

**Ready Issue**:
A GitHub Issue that a human has labelled `ready-for-agent`. Only a Ready Issue may enter the Closed Loop.
_Avoid_: ticket, task

**Run Log**:
The structured record that the released app writes for each batch run, with no personal patient data.
_Avoid_: debug log, telemetry
