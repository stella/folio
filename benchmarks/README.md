# Folio performance benchmarks

Folio has two complementary benchmark layers:

- `bun run bench` measures isolated parse, serialize, and Markdown operations.
- `bun --filter @stll/folio-core perf` builds the React playground for production
  and measures browser-visible open and editing workflows.

The browser profiler is the source of truth for user-facing latency. A faster
isolated function is not an improvement unless the corresponding end-to-end
scenario also improves.

## Browser baseline

Install dependencies and Chromium once:

```sh
bun install
bunx playwright install chromium
```

Run the production baseline:

```sh
bun --filter @stll/folio-core perf
```

The command writes unprefixed, machine-readable JSON to
`.cache/performance/folio-perf.json`. It also emits the JSON on standard output;
when invoked through Bun's workspace filter, Bun prefixes console lines, so use
the artifact for automation.

Useful controls:

```sh
# Collect enough samples to study variance before proposing budgets.
FOLIO_PERF_REPETITIONS=20 bun --filter @stll/folio-core perf

# Run one fixture quickly, or the complete documented performance corpus.
FOLIO_PERF_SUITE=smoke bun --filter @stll/folio-core perf
FOLIO_PERF_SUITE=corpus bun --filter @stll/folio-core perf

# Enable method wrappers to diagnose call counts and their wall-time overhead.
FOLIO_PERF_DIAGNOSTICS=1 bun --filter @stll/folio-core perf

# Reuse an already built playground or an externally managed production URL.
FOLIO_PERF_SKIP_BUILD=1 bun --filter @stll/folio-core perf
FOLIO_PERF_URL=http://127.0.0.1:4201 bun --filter @stll/folio-core perf

# Write elsewhere; relative paths are resolved from the repository root.
FOLIO_PERF_OUTPUT=.cache/performance/noise-study.json \
  bun --filter @stll/folio-core perf
```

Do not use Vite development-mode Playwright timings as the baseline. The
`performance` Playwright project remains valuable for structural invariants such
as incremental measurement and page virtualization. Give each worktree a unique
port so `reuseExistingServer` cannot attach to another checkout:

```sh
FOLIO_PLAYGROUND_PORT=4300 bunx playwright test --project=performance
```

### Cache states

- `cold`: each sample uses a new browser context, so HTTP and browser caches are
  empty. Browser-process startup is excluded.
- `warm`: the scenario is primed once, then every measured sample uses a new page
  in the same browser context. Application and fixture responses may be cached;
  editor state is not reused.

Cold and warm samples are never aggregated together.

### Milestones

All browser milestones use `performance.now()` relative to the start of that
sample:

| Milestone               | Meaning                                                                                                                                          |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| `bytes-available`       | The fixture response has completed. Pre-parsed generated models omit this milestone.                                                             |
| `docx-parsed`           | `parseDocx` returned a document model.                                                                                                           |
| `prosemirror-ready`     | Document-model to ProseMirror conversion completed.                                                                                              |
| `editor-state-ready`    | ProseMirror state and its editor plugins were initialized. The DOM view remains intentionally deferred until interaction.                        |
| `fonts-ready`           | The initial layout font wait completed.                                                                                                          |
| `layout-start`          | The layout pipeline accepted an initial, transaction, font, input, or manual run. For edits, this separates input scheduling from layout work.   |
| `flow-blocks-ready`     | ProseMirror content was converted to layout flow blocks.                                                                                         |
| `measurement-ready`     | Block measurement completed.                                                                                                                     |
| `pagination-ready`      | Pagination completed.                                                                                                                            |
| `render-pages-ready`    | The page painter finished its synchronous work.                                                                                                  |
| `visible-pages-painted` | Visible page content survived two animation frames.                                                                                              |
| `first-usable-page`     | Painted document content and the editor's interaction controller are both available. The deferred ProseMirror view is not forced into existence. |
| `complete-layout`       | The layout pipeline completed pagination and page rendering. Virtualized off-screen page content need not be mounted.                            |

The report retains raw samples and provides median and p95 distributions for
wall time, milestones, browser CPU categories, heap, phase durations, long
tasks, page counts, DOM elements, and measured blocks. With fewer than 20
samples, treat p95 as directional; with three samples it is effectively the
maximum. Environment metadata includes the logical CPU count and host load
average before and after the run; discard or repeat runs affected by unrelated
host contention.

Standard mode is the default for wall-time results. Detailed mode wraps
`measureText`, `getBoundingClientRect`, and `createElement` to count and time
calls. Compare otherwise identical detailed and standard runs before using the
detailed timings; the report records its instrumentation mode.

## Microbenchmarks

```sh
bun run bench
```

Tinybench warms every task before measurement and executes repeated samples.
The report includes median and p95 latency. Microbenchmarks deliberately copy
input bytes for each parse so retained or transferred buffers cannot influence
subsequent iterations.

## Corpus

Only repository-owned, synthetic, or upstream-redistributable fixtures may be
used. Never add customer or other private documents.

| Input                                                               | Purpose                                                                                                      | Provenance                                                                                                                                             |
| ------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `tests/visual/fixtures/docx-editor-demo.docx`                       | Small ordinary document; lists, tables, headers/footers, notes, and comments.                                | Imported with the Eigenpal `docx-editor` fork under the MIT attribution preserved in `NOTICE.md`.                                                      |
| `tests/visual/fixtures/sample.docx`                                 | Medium mixed-content parse and serialization case.                                                           | Imported with the Eigenpal `docx-editor` fork under the MIT attribution preserved in `NOTICE.md`.                                                      |
| `tests/visual/fixtures/podily-bps.docx`                             | Larger rich document with lists, tables, headers/footers, images, and many paragraph identifiers.            | Imported with the Eigenpal `docx-editor` fork under the MIT attribution preserved in `NOTICE.md`.                                                      |
| `tests/visual/fixtures/performance-1500-paragraphs.docx`            | Large end-to-end parse, conversion, layout, and paint case.                                                  | Deterministically generated by `benchmarks/build-performance-fixtures.ts` from repository-authored text.                                               |
| `tests/visual/fixtures/performance-long-split-table.docx`           | A 121-row table whose first row is taller than a page, exercising row fragmentation and repeated pagination. | Deterministically generated by `benchmarks/build-performance-fixtures.ts` from repository-authored text.                                               |
| `tests/visual/fixtures/performance-mixed-script-embedded-font.docx` | RTL Arabic/Hebrew, CJK, Latin, and mixed-script paragraphs plus an embedded font.                            | Deterministically generated by `benchmarks/build-performance-fixtures.ts`; the embedded Arimo Hebrew face is Apache-2.0 data from `@fontsource/arimo`. |
| `?paragraphs=1500`                                                  | Deterministic pre-parsed model used only for incremental editing, where parsing would be setup noise.        | Generated in `packages/playground/src/App.tsx` from repository-authored text.                                                                          |
| `packages/core/src/docx/__tests__/__fixtures__/corpus/*.docx`       | Focused OOXML compatibility and round-trip cases.                                                            | Deterministically generated by `packages/core/scripts/build-corpus-fixtures.ts`; see the adjacent `PROVENANCE.md`.                                     |

Together, the corpus suite covers an ordinary small file; the demo and rich
fixtures; a real 1,500-paragraph DOCX; split table rows; tracked changes and
comments; headers, footers, footnotes, and endnotes; an image and embedded font;
RTL, Arabic/Hebrew, CJK, and mixed-script content; and packages with many styles,
relationships, and parts. Regenerate synthetic files with:

```sh
bun benchmarks/build-performance-fixtures.ts
```

## Comparing changes

Use the same commit build settings, browser version, hardware, instrumentation
mode, cache state, viewport, fixture, and repetition count. Report both median
and p95, include observed variance, and retain correctness gates. Do not commit
one developer machine's JSON as a universal baseline or enforce a budget until
the measurement noise is understood.
