# Differential parser testing (folio vs reference parsers)

This directory holds a small scaffold for cross-checking folio's DOCX
parser against established reference parsers. PR #587 (block-level
content controls) surfaced a handful of prefix / namespace / OnOff edge
cases where a spec-faithful implementation still disagreed with how
mature parsers actually interpret the wire format. Differential testing
catches that class of issue directly.

The harness projects folio's parse of every corpus fixture
(`__tests__/__fixtures__` plus the repo's visual fixtures) into a structural
shape and asserts it matches the same projection taken from each reference
parser. References today:

- **python-docx** — community Python wrapper over `lxml`
- **Open XML SDK** — Microsoft's MIT-licensed .NET OOXML library
  (`DocumentFormat.OpenXml`)

Locally the suite skips when a reference is missing so `bun test` stays
runnable; CI installs both and sets `DIFFERENTIAL_REQUIRED=1`, which turns a
missing dependency into a failure so the parity gate cannot silently pass
(`bun run test:differential`). Property-based generators are a follow-up.

## Why these references (and not docx4j)

**python-docx**

- **No JVM.** Runs in well under a second per small fixture; trivial to
  shell out from a Bun script.
- **Already on the dev image.** Python 3.x is preinstalled on macOS dev
  machines and most CI runners; the only extra step is
  `pip install python-docx`.
- **Wire-format-friendly API.** python-docx exposes the underlying
  `lxml` element tree (`document.element.body`), so the projector can
  count `w:p` / `w:r` / `w:tbl` / `w:sdt` directly against the OOXML
  XPath — closer to "what the file actually says" than the high-level
  `paragraphs`/`tables` API.

**Open XML SDK**

- **Microsoft-adjacent reference.** MIT-licensed, maintained by the .NET
  Foundation; useful as a second opinion without replacing python-docx.
- **Lighter CI than docx4j.** `setup-dotnet` plus a single NuGet package;
  no JVM.
- **Dev/CI only.** The projector is not shipped in `@stll/folio-core`.

docx4j was considered as an alternative but it requires a JVM in CI and
a much larger install footprint. We can revisit it later if there is a
specific behaviour neither reference can model (e.g., advanced field
resolution).

## Setup

```bash
# python-docx (one-time, on dev or CI image):
pip install --user python-docx

# Open XML SDK projector (one-time, requires .NET 8 SDK):
dotnet build packages/core/scripts/differential/dotnet -c Release
```

The differential test (`packages/core/src/docx/__tests__/differential.test.ts`)
auto-skips a reference when its toolchain is missing, so contributors without
python-docx or .NET see a clean local test run rather than a hard failure.

## Running

Against a single DOCX:

```bash
# default reference: python-docx
bun packages/core/scripts/differential/diff.ts path/to/file.docx

# explicit reference
bun packages/core/scripts/differential/diff.ts path/to/file.docx open-xml-sdk
bun packages/core/scripts/differential/diff.ts path/to/file.docx python-docx
```

Exit codes:

- `0` — projections are equivalent.
- `1` — at least one structural divergence (printed to stderr).
- `2` — infrastructure error (reference missing, fixture missing, parse
  exception).

The test suite runs the same harness against the full corpus for each
available reference.

## Projection shape

All projectors emit the same normalised JSON shape (see
`StructuralProjection` in `projection.ts`):

- `totalParagraphs` — all `w:p` elements reachable from the body
  subtree, including paragraphs nested in tables and block SDTs.
- `totalTables` — all `w:tbl` elements reachable from the body subtree.
- `topLevelBlocks` — direct paragraph/table/SDT children of `w:body`.
- `sdts[]` — `{ scope, sdtType, alias?, tag?, lock?, childCount }` for
  each `w:sdt` in document order.
- `sdtCountsByType` — quick eyeball summary of the SDT inventory.

### What we do not project (and why)

- **Run count.** folio applies a run consolidator: adjacent
  identically-formatted `w:r` elements collapse into one. Wire-format
  run counts will therefore always diverge from python-docx on real
  documents. Adding `totalRuns` back would make every interesting
  fixture diverge for an uninteresting reason. If run-level parity
  becomes interesting, compare _consolidated_ runs on both sides
  instead of raw `w:r`.
- **Split inline SDT segments are coalesced.** Folio splits a wire
  `w:sdt` whose inline content straddles a lifted marker (bookmark /
  comment range / tracked-change boundary) into multiple `InlineSdt`
  segments that share one `SdtProperties` reference. The folio
  projector merges those segments back into a single entry (summing
  `childCount`) so it matches the reference projector's one-entry-per-
  `w:sdt` view; otherwise every bookmarked or comment-marked control
  would diverge on `sdts` length for an uninteresting reason.
- **Complex field runs are collapsed.** folio represents a `w:fldChar`
  begin/separate/end span (plus its `w:instrText` and result runs) as a
  single `complexField` `ParagraphContent` item, so an inline SDT
  containing a PAGE / REF / etc. field reports `childCount: 1` for the
  whole field. The reference projectors apply the same collapse rule when
  computing `childCount` for inline SDTs: every `w:r` from a `begin`
  fldChar through the matching `end` fldChar counts as one logical
  child. Nested complex fields (e.g. PAGEREF inside a TOC entry) are
  tracked with a depth counter so an inner pair does not close the outer
  span early. Without this rule any fixture with a complex field inside
  an inline content control would diverge on `childCount` even when
  folio's parse is structurally correct.
- **Textbox content.** folio models drawing-anchored paragraphs and
  tables as run-level shape content, not block content. The reference
  projectors exclude paragraphs/tables/SDTs inside `w:txbxContent` to
  keep the body-block comparison apples-to-apples. Textbox parity is
  tracked as a separate concern.

## What the references cover (and don't)

Covers cleanly:

- Block structure (`w:p`, `w:tbl`, `w:sdt`) at any depth.
- SDT properties recoverable from `w:sdtPr` (alias, tag, lock, type).
- Headers/footers/footnotes (via related parts — not used yet by this
  projection; would extend cleanly).

Does not cover:

- Run consolidation semantics (references preserve wire-format runs as-is).
- Anything below `w:r` (text content, breaks, tabs) without bespoke
  walking — fine for our needs, but would expand if we projected
  run-level content.
- Advanced field/complex-field resolution.
- Drawing/SmartArt internals.

## Files

- `projection.ts` — folio-side structural projection.
- `python_docx_project.py` — python-docx-side structural projection.
- `dotnet/` — Open XML SDK projector (`OpenXmlProjector`, .NET 8).
- `diff.ts` — orchestrator CLI; exported `runDifferential` is reused by
  the differential test.
- `../../src/docx/__tests__/differential.test.ts` — full-corpus parity
  test for each reference.
