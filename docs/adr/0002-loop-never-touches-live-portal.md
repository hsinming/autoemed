# The Closed Loop never touches the Live Portal

Automated tests mock helium and never connect to the eMedical Live Portal, because the portal holds real health cases and an agent that clicks "Normal" or saves would change a real applicant's record. The only contact with the Live Portal is the Smoke Test: a human starts it, completes 2FA, types one eMed No., and the test runs read-only steps (search, open Manage Case, check selectors, Close) and stops before any field is clicked.

## Consequences

- A change to `EMEDICAL_URL`, a selector or the portal flow cannot be proven by the Quality Gates alone; the loop stops with a Handoff and asks a human to run the Smoke Test.
- No eMed No. is stored in the repo for testing.
