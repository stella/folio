# Durable block IDs in DOCX

Goal: every block in every DOCX is addressable by a stable ID that survives
editing in folio and identity-preserving DOCX round-trips, so the positional
`seq-NNNN` fallback (known bug class: sequential IDs renumber after structural
edits) becomes unreachable in practice. This unlocks per-block blame, robust
comment and citation anchors, and AI-edit targeting for host applications.

Provenance note: the direction was inspired by ideas in macro-inc/macro
(AGPL-3.0). Ideas only; no code was read or ported. Everything here is
original work against the OOXML spec and the existing folio code.

## Status

| Work item                                   | Status                                                                                                                                                 |
| ------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------ |
| WI-1 ingest normalization (`ensureParaIds`) | Implemented — `packages/core/src/docx/ensureParaIds.ts`, exported from `@stll/folio-core/server`                                                       |
| WI-2 editor ID lifecycle spec + tests       | Implemented — `packages/core/src/prosemirror/extensions/features/ParaIdAllocatorExtension.ts` (header documents the rules; co-located tests lock them) |
| Table identity                              | Decided against durable table-level IDs — see below                                                                                                    |
| WI-5 touched-paraIds change reporting       | Design only — see below                                                                                                                                |

## What already existed (do not rebuild)

| Piece                                                                              | Where                                                                                                                        |
| ---------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `w14:paraId` / `w14:textId` parsed off `<w:p>` (with a `w:paraId` prefix fallback) | `packages/core/src/docx/paragraphParser.ts`                                                                                  |
| Both serialized back, emitted only when truthy                                     | `packages/core/src/docx/serializer/paragraphSerializer.ts`                                                                   |
| In-editor allocator backfilling null/duplicate paraIds per doc-changed transaction | `packages/core/src/prosemirror/extensions/features/ParaIdAllocatorExtension.ts`                                              |
| Hex ID generation with the OOXML `ST_LongHexNumber` bound                          | `packages/core/src/utils/hexId.ts`                                                                                           |
| Deterministic load-time backfill in the headless reviewer                          | `packages/core/src/ai-edits/headless.ts` (`ensureDeterministicParaIdsInState`)                                               |
| Live lookup by paraId                                                              | `packages/core/src/prosemirror/utils/findParagraphByParaId.ts`                                                               |
| Live-first block resolution, snapshot fallback for `seq-`                          | `packages/react/src/components/aiEditRange.ts` (`resolveFolioAIBlockRange`; `scrollToBlock` in `DocxEditor.tsx` is a caller) |
| String-offset part patching without the PM model                                   | `packages/core/src/docx/selectiveXmlPatch.ts`                                                                                |
| Deterministic comment-part paraIds at save                                         | `packages/core/src/docx/serializer/commentSerializer.ts`                                                                     |

## WI-1: `ensureParaIds` (implemented)

`ensureParaIds(docx: Uint8Array | ArrayBuffer) -> Promise<{ docx: Uint8Array,
assigned, deduplicated, alreadyComplete }>`. Hosts call it once at
ingest/upload, before deriving or storing any block anchors, so every stored
working version has full ID coverage. Word 2010+ writes
paraIds; Google Docs exports, LibreOffice, and python-docx/docx4j output
generally do not — those are exactly the documents that fell into `seq-`.

Contract highlights (the module doc in `ensureParaIds.ts` is normative):

- Patches part XML in place with string splices (no model round-trip), so
  unmodeled features survive byte-identical. Covers `word/document.xml`,
  headers, footers, footnotes, endnotes; table-cell and `mc:Choice` text-box
  paragraphs are nested `<w:p>` in those parts and are covered by the same
  scan. The comments part mints its own deterministic paraIds at save time
  and is left untouched (its IDs still count toward uniqueness).
- `mc:Fallback` content is never modified: Word regenerates the fallback
  branch (duplicating the `mc:Choice` IDs) on save, so stamping there would
  churn on every Word round-trip.
- Fresh IDs are deterministic (`deterministicHexId` over a document
  fingerprint, part path, and paragraph ordinal): the pass is a pure function
  of the input bytes, so retried ingests produce identical output.
- Deviation from the original proposal: later duplicates ARE reassigned
  (first occurrence keeps the ID). The proposal said "existing IDs are never
  rewritten", but that conflicts with its own zero-`seq-` acceptance goal,
  and duplicated IDs already break `findParagraphOffsets` (ambiguous → null)
  and the editor allocator rewrites them on first edit anyway. Reassigning at
  ingest matches what the document becomes after one editor session.
- The reserved all-zero paraId is treated as unassigned (Word semantics).
  `generateHexId` / `deterministicHexId` now exclude `00000000` at the
  source.
- Patched parts get `xmlns:w14`, `xmlns:mc`, and an `mc:Ignorable` listing
  `w14` injected on the root when absent; non-Word producers declare none of
  the three.
- `w14:textId` is written alongside a newly minted paraId (same value),
  never touched otherwise. It is a text-revision marker, not identity; never
  key anything on it.
- Idempotent: full-coverage input returns the original bytes
  (`alreadyComplete: true`), byte-identical by construction. (A no-op rezip
  would NOT be byte-identical — JSZip regenerates the container — which is
  why the short-circuit exists.)
- OPC digital signatures cover package bytes and become invalid when a signed
  package is rewritten. A signed document that is already complete is returned
  untouched. If normalization is required, the default is to reject it;
  callers may set `allowSignedPackageMutation` only after warning the user.

### Microsoft Word round-trip result

A real Word edit/save check established the boundary of the durability
contract:

- When a Word-saved package had one paragraph ID pair removed,
  `ensureParaIds` restored it and Word preserved all 26 paragraph IDs on the
  next edit/save. This is the normal identity-preserving round-trip.
- On the first Word save of a non-Word-produced package, Word replaced
  `w14:docId` and every paragraph ID. Supplying a valid, previously unseen
  `w15:docId` did not change that behavior. IDs minted before Word adopts the
  package therefore do not survive that first save.

The second case is outside this function's control: the IDs are valid OOXML,
but Word elects to establish a new document identity. Hosts that accept an
externally edited package must normalize it again; if anchors must cross that
first Word save, they need content-based reconciliation rather than assuming
the old and new paraIds match.

## WI-2: editor ID lifecycle (implemented)

The rules are documented in the `ParaIdAllocatorExtension` header and locked
by its tests:

- Split: the half holding the paragraph's original start position keeps the
  ID (anchors living in the content stay resolvable); the other half gets a
  fresh one.
- Merge/join: the survivor keeps its ID; the absorbed ID dangles by design
  (consumers degrade to snapshot fallback).
- Paste/duplicate of an ID already in the doc: resolved by mapping the
  original paragraph's position through the transaction, so pasting a copy
  ABOVE its source no longer steals the source's ID (a real bug under the
  previous first-in-document-order rule). Fallback when nothing maps back:
  first occurrence keeps.
- Paste of an ID unknown to the doc (cross-doc paste, cut-then-paste move):
  kept, so moving a paragraph preserves its anchors.
- Undo/redo: allocation applies with `addToHistory: false`; ProseMirror remaps
  the allocation into the originating history event, so redoing a split
  restores the same allocated ID.
- The appended allocation transaction is excluded from paragraph change
  tracking (`ignoreTrackedChanges`), matching `ensureParaIdsInState`;
  backfilling an untouched paragraph must not mark it as user-changed.

Both changes are deliberate divergences from the eigenpal upstream; the file
header notes them.

## Table identity: decided against durable table-level IDs

The framing document called out that tables/rows/cells carry no identity
(`TableExtension.ts` node specs have no id attrs) but lost the work items
covering it. Decision: do not add them. OOXML has no `paraId` equivalent for
`<w:tbl>` / `<w:tr>` / `<w:tc>`, and Word strips unknown attributes on save,
so any invented table ID would be durable only inside folio — a false sense
of stability that dies on the first Word round-trip, which is exactly the
failure mode this effort exists to remove. Tables are addressed indirectly
through the paraIds of their contained paragraphs, which `ensureParaIds`
guarantees exist (including in cells). If a session-scoped table identity is
ever needed for UI purposes, it must be clearly non-persistent.

## WI-5: touched-paraIds change reporting (design only)

For downstream per-block blame (host-side, keyed by entity + paraId + author

- timestamp), folio needs to report which blocks a change touched.
  `applyAIEditOperations` / `applyDocumentOperations` return `{ applied,
skipped }` where each applied entry carries `revisionIds` (note: nested per
  entry, not top-level). Design: extend each applied entry with the affected
  `blockIds` (paragraph-granular), and add a ref method returning paraIds
  touched since a given snapshot. Implement when the host blame schema lands;
  do not ship the surface speculatively.

## Constraints and gotchas

- `w14:paraId` validity: 8 uppercase hex chars, not `00000000`, below
  `0x80000000`, unique document-wide. `hexId.ts` constants are the single
  source of truth; reuse, don't reimplement.
- Durability follows the document identity. Microsoft Word can replace
  `w14:docId` and all paraIds on its first save of a non-Word-produced package;
  never promise pre-save paraId continuity across that boundary.
- `deriveBlockId` (`packages/core/src/types/block-id.ts`) returns the paraId
  VERBATIM when present and non-duplicate; it does not mint hex. Its
  contract is published (host server extractors and prompts consume it);
  `seq-` support must remain for legacy snapshots. The goal is that new
  snapshots never contain it.
- `w:rsid*` attributes on `<w:p>` are dropped on import. Out of scope; leave
  as-is.
- The serializer emits `w14:paraId` only when truthy; after ingest
  normalization plus the allocator that is always, but keep the guard.
- `ParaIdAllocatorExtension` and `hexId.ts` carry upstream-sync headers for
  the eigenpal fork; divergences must be noted there (both now are).

## Host-side follow-ups (separate work, not in this repo)

Calling `ensureParaIds` at upload/ingest and collaborative finalize; the
blame storage and UI; migrating any paragraph-index-based redline tooling
onto paraId addressing.
