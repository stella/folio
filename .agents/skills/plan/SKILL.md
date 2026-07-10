---
name: plan
description: "Create a new implementation plan in the repo's planning area."
---

# Create Plan

Create a new implementation plan in the repo's planning area.

## Arguments

$ARGUMENTS — A short slug for the plan, ideally kebab-case. If empty,
derive a reasonable slug from the task.

## Conventions

Prefer this shared planning layout when the repo supports it:

```text
.agents/ARCHITECTURE.md
.agents/GOALS.md
.agents/STATUS.md
.agents/plans/
```

If the repo uses a different planning area, adapt to it rather than
creating duplicate systems.

## Instructions

1. **Read planning context**:
   - `.agents/ARCHITECTURE.md` if present
   - `.agents/GOALS.md` if present
   - recent related plans if present

2. **Determine the plan location**:
   - prefer `.agents/plans/`
   - if that directory does not exist, create it only if the repo is
     clearly adopting the shared planning convention
   - otherwise ask or use the repo's existing planning area

3. **Determine the next plan number**:

   ```bash
   ls .agents/plans/
   ```

   Use the next sequential number when numbered plans are already in use.
   If the repo does not number plans, follow its established naming.

4. **Research before writing**:
   - inspect the existing code
   - read the relevant modules
   - understand constraints, patterns, and likely blast radius

5. **Write the plan** with this structure:

   ```markdown
   # Plan: [Feature Name]

   Date: YYYY-MM-DD

   ## Goal

   What are we building and why?

   ## Design Decisions

   - **Decision**: Why this approach over alternatives.

   ## Scope

   **In scope:**

   - ...

   **Out of scope:**

   - ...

   ## Implementation

   - `path/to/file` — what changes here

   ## Test Cases

   What needs to be verified.

   ## Open Questions

   Any unresolved decisions.
   ```

6. **Plan well, do not over-specify**:
   - focus on _what_ and _why_
   - avoid locking in incidental implementation details too early
   - call out migrations, security implications, and rollout risk when relevant

7. **Confirm before finalizing** if the user is still shaping the task.
   If they asked you to just create the plan, go ahead and write it.
