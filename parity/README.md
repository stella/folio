# Word rendering parity engine

Answers one question for any `.docx`: does folio lay it out the way real
Microsoft Word does, and if not, where and why?

Like `benchmarks/cross-language/`, this is an on-demand local tool, not a CI
gate: the ground truth requires Microsoft Word for Mac (driven via
AppleScript) plus `mutool` (`brew install mupdf-tools`). The tool skips with a
clear message when either is missing.

## Usage

```sh
bun parity/cli.ts                          # full default corpus, HTML report
bun parity/cli.ts path/to/file.docx        # one document
bun parity/cli.ts some/dir --json          # machine-readable output
bun parity/cli.ts --refresh-truth          # ignore cached Word exports
```

The HTML report lands in `parity/report/index.html`.

For repeated runs in one worktree, start a current-source playground once and
explicitly reuse it:

```sh
FOLIO_PLAYGROUND_PORT=$(bun -e 'import { PLAYGROUND_PORT } from "./parity/config"; process.stdout.write(String(PLAYGROUND_PORT))') \
  bun --filter @stll/playground dev
bun parity/cli.ts path/to/file.docx --reuse-server
```

Fresh-server mode remains the default. Reuse is an opt-in for iterative work
where the caller owns the server and knows it reflects the current worktree.

## How it works

1. **Ground truth** (`wordTruth.ts`): the docx is exported to PDF by scripting
   the locally installed Word, then `mutool draw -F stext` extracts per-line
   text + geometry (points, page-relative). Artifacts are cached in
   `parity/.cache/<sha256>/` so re-runs never reopen Word.
2. **Folio extraction** (`folioExtract.ts`): the same docx is loaded in the
   playground (Playwright + the existing `?file=` fixture route) and the
   painted DOM (`.layout-page` / `.layout-line`) is walked into the same
   normalized geometry, converting CSS px to pt and normalising against the
   page element width so playground zoom cannot skew coordinates. Per-page
   screenshots are captured for the report.
3. **Comparison** (`compare.ts`): lines are sequence-aligned by text, which
   surfaces pagination shifts, line-break differences, and missing/extra
   content; matched lines are then diffed geometrically with a per-page
   median-offset correction (Word boxes are ink bounds, folio boxes are
   line-height boxes, so only residual drift is a divergence).
4. **Feature attribution + clustering** (`features.ts`): each paragraph in the
   OOXML is tagged with the layout features it exercises (tables, tabs,
   numbering, anchored drawings, spacing rules, justification, fields, CJK,
   ...); each divergence inherits the tags of its paragraph. Across the corpus,
   (divergence kind x feature) clusters are ranked by lift, so a single root
   cause (e.g. "atLeast line spacing") shows up as one ranked cluster instead
   of dozens of scattered diffs.
5. **Report** (`report.ts`): per-doc fidelity scores, the cluster ranking, and
   per-page Word-vs-folio renderings with both sets of line boxes overlaid.

Font preflight compares the family Word embedded in its PDF with Chromium's
resolved CSS family on matching text lines. A requested font may be substituted
without invalidating the run when both renderers use the same fallback
(`font-shared:*`). `font-renderer-mismatch` means geometry may primarily reflect
different font metrics and the document should not drive a layout fix.
Chromium families are canvas-probed in CSS-stack order; the computed stack alone
cannot reveal which fallback supplied the glyphs.

Shared data contract: `types.ts`. Tolerances and paths: `config.ts`.
