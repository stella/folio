# DOCX rendering interoperability harness

Answers one question for any `.docx`: how does folio's layout compare with an
independent DOCX renderer, and where do the implementations diverge?

No renderer is labelled “ground truth.” The default reference is open-source
[LibreOffice Writer](https://www.libreoffice.org/). Microsoft Word is an
explicit, optional reference. Agreement with either implementation is useful
interoperability evidence, but it does not by itself prove OOXML conformance.

This is an on-demand local tool, not a CI gate. It uses `mutool`
(`brew install mupdf-tools`) to normalize each exported PDF into the same line
geometry.

## Usage

```sh
# LibreOffice is the default reference renderer
bun parity/cli.ts
bun parity/cli.ts path/to/file.docx --reference libreoffice

# Word is available only when selected explicitly
bun parity/cli.ts path/to/file.docx --reference word

bun parity/cli.ts some/dir --json
bun parity/cli.ts --refresh-reference
```

Requirements:

- `libreoffice`: a local LibreOffice installation plus `mutool`; the adapter
  runs Writer headlessly with an isolated temporary profile.
- `word`: Microsoft Word for Mac plus `mutool`; the adapter drives a locally
  installed copy through AppleScript. Check the licence terms covering that
  installation before using automated outputs for compatibility development.

The HTML report lands in `parity/report/index.html` and identifies the selected
renderer and its version. Reports describe the renderer as a **reference**, not
as a specification oracle.

## Word line-endpoint regression manifests

The focused line-endpoint validator separates Word capture from Folio
validation. Capture requires the local Word adapter, but validation reads a
versioned JSON manifest and does not launch or require Word. This makes a
reviewed reference capture reusable on machines and CI runners where Word is
not installed.

```sh
# On a licensed Mac with Microsoft Word and mutool
bun run parity:line-endpoints capture path/to/fixture.docx \
  --output path/to/fixture.word-lines.json

# Anywhere Folio's playground can run; Word is not opened
bun run parity:line-endpoints validate path/to/fixture.docx \
  --manifest path/to/fixture.word-lines.json
```

The repository baseline for automatic hyphenation and hanging punctuation is
fully synthetic and reproducible:

```sh
bun run parity:build-line-endpoint-fixtures
bun run parity:line-endpoints capture \
  parity/fixtures/word-hyphenation-hanging.docx \
  --output parity/fixtures/word-hyphenation-hanging.word-lines.json \
  --refresh-word
bun run parity:line-endpoints validate \
  parity/fixtures/word-hyphenation-hanging.docx \
  --manifest parity/fixtures/word-hyphenation-hanging.word-lines.json
```

Rebuilding the DOCX is deterministic. Recapturing the manifest is not a
routine snapshot update: inspect the line texts, record the local Word and
`mutool` versions already embedded in the manifest, and confirm the change is
an intended reference-behavior update before committing it.

The manifest records the exact DOCX SHA-256, Word and extraction versions, and
the normalized text occupying each visual line. Validation fails fast if the
DOCX hash differs, then reports only page placement and line-ending
differences; x/y/width geometry is intentionally excluded. Because normalized
line text can still contain sensitive document content, do not commit manifests
captured from confidential or personal documents. Use synthetic or explicitly
public fixtures for repository baselines.

This is a clean-room interoperability check: Word runs only as an external
reference renderer during capture. No Word code, APIs, or runtime dependency
ships in Folio. A captured result is evidence about one documented Word version
and font environment, not proof of OOXML conformance.

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

1. **Reference rendering**: the selected adapter exports the DOCX to PDF.
   LibreOffice uses a headless isolated process; Word uses its macOS scripting
   interface. `mutool draw -F stext` extracts per-line text and geometry.
   Artifacts are cached by document hash and renderer, so repeat runs do not
   reopen the application.
2. **Folio extraction** (`folioExtract.ts`): the same DOCX is loaded in the
   playground and the painted DOM is walked into normalized geometry. Page
   screenshots are captured for the report.
3. **Comparison** (`compare.ts`): lines are sequence-aligned by text, exposing
   pagination shifts, line-break differences, and missing or extra content.
   Matched lines are then compared geometrically.
4. **Feature attribution and clustering** (`features.ts`): each divergence is
   associated with the active OOXML features. Cross-corpus clusters expose a
   likely shared cause instead of presenting dozens of isolated diffs.
5. **Report** (`report.ts`): per-document scores, cluster rankings, and
   side-by-side reference/Folio renderings with line-box overlays.

Font preflight compares the family embedded in the reference PDF with
Chromium's resolved CSS family. A requested font may be substituted without
invalidating the run when both renderers use the same fallback. A
`font-renderer-mismatch` means geometry may primarily reflect different font
metrics and should not drive a layout change by itself.

## Interpreting disagreements

The engineering source order is:

1. [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
   and ISO/IEC 29500 normative requirements;
2. published implementation notes and documented extensions;
3. structural differential checks against python-docx and Open XML SDK;
4. reproducible behavior across independent renderers.

When LibreOffice, Word, and folio disagree, the harness records the divergence;
it does not automatically crown the majority. Page layout contains
underspecified and implementation-specific behavior, and different engines may
also substitute different fonts. A fix should be tied to a specification rule,
documented extension, or an explicit interoperability choice.

See [Interoperability references](../docs/interoperability.md) for the projects
we can legally and technically use, what each one can measure, and why some
DOCX libraries are not layout references.

Shared data contract: `types.ts`. Tolerances and paths: `config.ts`.
