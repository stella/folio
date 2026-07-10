---
name: rabbit-round
description: "Process automated PR review comments systematically. Use this for CodeRabbit, Gemini, GitHub Copilot, Devin, Greptile, and similar bots."
---

# Rabbit Round

Process automated PR review comments systematically. Use this for
CodeRabbit, Gemini, GitHub Copilot, Devin, Greptile, and similar bots.

## Instructions

1. **Get context**:

   ```bash
   PR_NUMBER=$(gh pr view --json number -q '.number')
   gh api user --jq '.login'
   git rev-parse HEAD
   ```

2. **Fetch review comments** from the PR:

   ```bash
   gh api repos/{owner}/{repo}/pulls/"$PR_NUMBER"/comments --paginate
   ```

   Filter for known bot accounts. Do not treat human review comments as bot comments.

3. **Triage each bot comment**:
   - **Accept** if it improves correctness, safety, maintainability, or follows repo conventions.
   - **Push back** if it is incorrect, overreaching, or conflicts with documented conventions.

4. **Implement accepted suggestions**:
   - make the code changes
   - group related fixes logically
   - run the relevant project checks

5. **Reply inline** to each bot comment:

   ```bash
   COMMENT_ID="123456789"
   gh api -X POST repos/{owner}/{repo}/pulls/"$PR_NUMBER"/comments/"$COMMENT_ID"/replies \
     -f body="[response]"
   ```

   Good response patterns:
   - accepted and implemented
   - agreed, implemented with a small adjustment
   - already addressed in commit `{hash}`
   - pushing back, with a concrete reason and source or repo convention

6. **Never resolve human review threads**. For bot threads, resolve only after:
   - the change is implemented, or
   - the pushback is clearly documented

7. **Check nit-level comments too**. Small ones still matter if they improve clarity
   or remove avoidable friction.

8. **Commit and push** if you made changes:
   - use a neutral commit message such as `fix: address review comments`
   - push to the active PR branch

## Decision Guidelines

**Accept when the suggestion:**

- fixes a bug or real edge case
- improves type safety
- adds missing tests
- aligns with existing repo patterns
- tightens security or validation appropriately

**Push back when the suggestion:**

- assumes facts not true in this codebase
- conflicts with canonical specs or official sources
- adds complexity for little benefit
- would undo a deliberate, documented decision
- is purely stylistic and inconsistent with the repo
