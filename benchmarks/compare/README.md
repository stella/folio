# Compare benchmark

Measures `compareDocx` on a generated corpus, and checks that what it produced
is worth measuring.

```sh
bun benchmarks/compare/run.ts                        # the whole generated corpus
bun benchmarks/compare/run.ts --sizes s,m --quick    # a fast loop while editing
bun benchmarks/compare/run.ts --filter tables        # one class, size or variant
bun benchmarks/compare/run.ts --out report.json      # the full machine-readable run
bun benchmarks/compare/run.ts --baseline             # record product digests
bun benchmarks/compare/run.ts --check                # prove an edit changed nothing
```

Recorded numbers live in `RESULTS.md`. Recorded digests live in `digests.json`.

## What it measures

Every configuration is a `class/size/variant` triple, run in its own process so
a large document cannot hand the next configuration a warmed JIT or a grown
heap. Within the process each sample is a whole comparison, bytes in to bytes
out: medians over nine samples after four discarded warm-ups, plus the heap
growth of one further un-timed run.

The four stages are timed separately: `parse` (both packages to editor models),
`align` (the pure planning pass), `apply` (writing tracked changes plus the
round-trip self-check), `serialize`. They are the functions `compareDocx`
itself composes, so the split cannot drift from the shipped pipeline.

### Document classes

| Class         | What it isolates                                                                |
| ------------- | ------------------------------------------------------------------------------- |
| `prose`       | The common path: headings and paragraphs, nothing else.                         |
| `lists`       | Multi-level numbering a redline must not silently renumber.                     |
| `tables`      | Merged, nested and empty cells: structure a flat block alignment cannot see.    |
| `notes`       | Footnotes and endnotes as separate stories with their own id space.             |
| `graphics`    | Images, equations and breaks: atoms a word diff must move whole or not at all.  |
| `fields`      | Simple and complex fields, hyperlinks and bookmarks spanning paragraphs.        |
| `sections`    | Headers, footers and multiple sections.                                         |
| `multiscript` | Right-to-left Arabic and Hebrew, CJK without spaces, and Latin interleaved.     |
| `revised`     | A base that already carries someone else's tracked changes.                     |

Sizes are block counts: `s` 40, `m` 320, `l` 2200.

### Edit variants

`identical` (the fixed cost of parse and serialize alone), `light` (one
paragraph in fifteen, the shape a review pass leaves), `heavy`, `churn`
(insertions, deletions and edits interleaved), `reorder` (a relocated run),
`structural` (splits, merges, a deleted row, a changed list level), `notes`
(the note stories only, main story byte-identical), `rewrite` (every
paragraph). A variant that would leave a class unchanged is skipped rather
than reported as a passing case.

Targets are built by rewriting the package XML, never by driving folio's own
applier: a target the engine produced is a target the engine agrees with by
construction, and generating it inside the measurement would put engine work
in the setup the stage timings exclude.

## What it asserts

Timings mean nothing without these, so a failing invariant fails the run.

- **`reject-returns-base` / `accept-returns-target`** — the round-trip algebra,
  stated as a comparison rather than a text equality: comparing the base with
  the rejected redline, and the target with the accepted one, must both report
  nothing. This runs the engine over its own output, which is where a redline
  that reads plausibly and is wrong shows up.
- **`self-compare-is-empty`** — a document compared with itself invents nothing.
- **`difference-is-reported`** — a pair the harness built to differ is reported
  as differing. Every other invariant here is satisfiable by seeing nothing, so
  without this one a blind spot passes: a change in a part the engine never
  reads survives accept and reject alike and self-compares clean.
- **`byte-determinism`** — two runs over the same inputs produce the same
  package bytes and the same change list.
- **`schema-validity`** — the redlined package passes the Open XML SDK
  validator. Skipped, loudly, when the machine has no .NET toolchain; build it
  with `dotnet build packages/core/scripts/differential/dotnet -c Release`.

`--baseline` records a SHA-256 of every product; `--check` reruns and reports
any that moved. That is how a performance change proves it altered nothing.

## An external corpus (local only)

`--corpus <dir>` adds pairs from a directory outside the repository. Nothing
from such a corpus is committed: third-party comparison corpora carry their own
licences, and a revision count is relative to the engine that produced it, so
they are a local measurement rather than a fixture. The suites simply do not
run without the flag.

The directory needs a `pairs.json`:

```json
[{ "id": "case-1", "base": "a.docx", "target": "b.docx", "expectedChanges": 4 }]
```

`expectedChanges` is optional and never asserted on; it is reported beside our
own count as a change detector. Failing a `pairs.json`, a directory of
`<case>/baseline.docx` plus `<case>/candidate.docx` is discovered
automatically.

`benchmarks/compare/.local/` is git-ignored, so it is a convenient place to
fetch one into.
