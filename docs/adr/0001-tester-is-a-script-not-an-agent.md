# The Tester is a deterministic script, and the commit hook reruns it

In the Closed Loop, no agent decides whether code passes. The Tester (`scripts/quality_gate.py`) runs syntax, lint and test gates and writes the Test Report; the `PreToolUse` hook on `git commit` reruns those gates live instead of trusting any saved report, and accepts the review only when `.loop/review.json` matches the staged tree hash. We chose this over an LLM tester because a Developer agent that can claim "done" will sometimes claim it falsely, and a stale or edited report must never let a commit through.

## Consequences

- Agents may add tests but may not delete or change existing test lines; the Tester fails the run if `git diff HEAD -- tests/` shows any removed line. Only a human changes an existing assertion.
- Human commits outside Claude Code are not gated (we used a Claude Code hook, not a git `pre-commit` hook).
