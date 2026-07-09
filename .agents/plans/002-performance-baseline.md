# Plan: Trustworthy Performance Baseline

Date: 2026-07-09

## Goal

Establish a reproducible production-build benchmark that separates DOCX load,
editor conversion, layout, and paint milestones; reports distribution statistics
and resource counters; and provides enough evidence to select the first focused
performance optimization.

## Design Decisions

- **Extend the existing profiler and layout instrumentation:** Preserve one
  benchmark path instead of creating competing harnesses.
- **Separate measurement from optimization:** This change records current
  behavior and measurement noise without changing editor algorithms or setting
  new end-to-end time budgets.
- **Emit versioned JSON:** Standard output is machine-readable; human diagnostics
  go to standard error, and raw samples remain available for independent analysis.
- **Measure production preview:** The runner builds the playground once and uses
  Vite preview, including the same repository-owned fixtures as development mode.
- **Name cache state explicitly:** Cold-cache and warm-cache samples use documented
  browser-context lifecycles and are aggregated independently.
- **Keep instrumentation optional:** Low-overhead wall-time runs and detailed
  call-count diagnostics can be compared so wrapper overhead is visible.
- **Use repository-owned inputs only:** Synthetic fixtures are deterministic and
  every checked-in fixture used by benchmarks has documented provenance.

## Scope

**In scope:**

- Repair the production profiler and preview fixture serving.
- Define versioned result, sample, milestone, environment, and aggregate types.
- Record parsing, ProseMirror conversion, font readiness, flow-block generation,
  measurement, pagination, rendering, first usable page, and complete layout.
- Record browser CPU categories, long tasks, heap, page and DOM counts, measured
  blocks, and relayout reasons.
- Run repeated cold-cache and warm-cache open scenarios plus representative local
  editing scenarios, reporting median and p95.
- Document commands, scenario semantics, corpus provenance, and interpretation.
- Preserve the existing Playwright performance tests as behavioral guardrails.

**Out of scope:**

- Performance optimizations or architectural rewrites.
- Rust or WebAssembly experiments.
- New aggressive CI timing thresholds.
- Claims that results from one developer machine are universal.
- Save, comparison, and large agent-operation benchmarks unless the common harness
  makes them trivial to add without widening the initial change substantially.

## Implementation

- `packages/playground/vite.config.ts` — serve safe DOCX fixture requests in both
  development and production preview.
- `packages/core/src/layout-engine/layoutInstrumentation.ts` — extend the optional
  instrumentation contract with load and user-visible milestones where ownership
  belongs in core.
- `packages/core/src/managers/DocumentLoaderManager.ts` — record DOCX parse timing
  at the parse boundary.
- `packages/core/scripts/profile-editor.ts` — run production scenarios repeatedly,
  collect CDP/browser metrics, aggregate samples, and emit versioned JSON.
- `tests/support/layoutMeasurement.ts` and performance tests — share updated event
  shapes and retain structural regression assertions.
- `benchmarks/README.md` — document reproduction, cold/warm semantics, output fields,
  instrumentation overhead, and how to collect traces.
- `benchmarks/fixtures.ts` or a corpus manifest — document each benchmark input's
  purpose and provenance.
- `.changeset/` — describe published-source changes or add an empty changeset when
  the instrumentation remains internal-only.

## Test Cases

- Production build and preview load every configured benchmark fixture.
- Profiler stdout parses as exactly one JSON document.
- Every sample contains ordered milestones and no layout errors.
- First usable page requires painted document content and an available editor view.
- Cold and warm results are distinct scenario groups with raw samples, median, and p95.
- A local edit reports a transaction relayout and does not remeasure the full
  1,500-paragraph document.
- Detailed instrumentation can be disabled without changing document behavior.
- Typecheck, unit tests, performance Playwright project, interaction tests,
  differential tests, and `git diff --check` pass proportionately.

## Open Questions

- How much wall-time overhead do the `measureText`, geometry, and DOM wrappers add
  on the benchmark host?
- Which repetition count produces a stable p95 without making the baseline suite
  impractically slow?
- Which open scenario contributes the first trace-supported optimization target?
- Should save and headless-operation scenarios extend this PR after the open/edit
  harness is stable, or land as the next measurement-only vertical PR?
