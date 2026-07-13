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
