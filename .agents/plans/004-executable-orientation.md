# Plan: Executable Developer Orientation

Date: 2026-09-09

## Goal

Turn Folio's existing architecture and testing rules into a queryable command that
maps a file or diff to its subsystem, pipeline seam, nearby tests, dependencies,
required checks, and release impact. A developer or agent should get a useful first
investigation route without reconstructing the repository from prose and search.

## Design Decisions

- **Derive most facts from the repository:** Discover packages, nearest AGENTS.md,
  adjacent tests, and TypeScript import edges from current files instead of copying
  them into a second maintained guide.
- **Keep policy small and typed:** Encode only the architectural seam and validation
  rules that cannot be inferred from paths or imports.
- **Support humans and automation:** Human output is concise; `--json` exposes a
  versioned result that later CI and agent tooling can consume.
- **Start from files and diffs:** Accept explicit paths and `--diff [ref]`, matching
  the two common entry points for regression work.
- **Trace source relationships, not runtime document data:** The first slice reports
  direct imports and importers. Runtime parse-to-paint traces remain a follow-up so
  this command stays fast and safe for arbitrary repository changes.

## Scope

**In scope:**

- Resolve explicit files, directories, or a Git diff to tracked repository files.
- Report package, seam, nearest instructions, adjacent tests, direct dependencies,
  direct importers, validation commands, and changeset requirements.
- Keep output deterministic and reject paths outside the repository.
- Add fixture-driven tests and dogfood the command on the layout and ZIP fixes.

**Out of scope:**

- Runtime DOCX content or layout tracing.
- Full transitive call graphs or symbol-level static analysis.
- Automatically executing suggested checks.
- Replacing architecture documents or package-level instructions.
- CI log summarization and the wider interaction-feature registry.

## Implementation

- `scripts/orient.ts` — CLI, repository discovery, import graph, policy mapping, and
  human/JSON output.
- `scripts/orient.test.ts` — path safety, diff/file discovery, seam classification,
  test adjacency, import tracing, and deterministic output.
- `package.json` — expose `bun run orient`.
- `AGENTS.md` source fragment — point future contributors to the command without
  duplicating its output.

## Test Cases

- A layout converter reports the flow-conversion seam, focused tests, core checks,
  and a published-package changeset requirement.
- A parity harness file reports parity checks without a changeset requirement.
- A directory expands deterministically and duplicate inputs collapse.
- `--diff` uses changed tracked files and rejects an invalid Git reference.
- Imports and importers resolve extensionless relative TypeScript paths.
- Paths outside the repository fail before content is read.
- JSON output is versioned, stable, and contains only repository-relative paths.

## Open Questions

- Should the next slice instrument parse-to-paint runtime provenance, or first add a
  machine-readable feature/interaction registry that can generate those trace hooks?
- Should CI consume the JSON to select checks, or remain advisory until the mapping
  has several weeks of observed accuracy?
