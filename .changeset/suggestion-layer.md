---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
"@stll/folio-agents": minor
"@stll/docx-core": minor
"@stll/folio-nuxt": patch
---

Add a first-class suggestion layer to tracked changes. AI-proposed edits can be
applied with the new `"suggested"` apply mode: they render with the
tracked-change grammar but a dotted stroke and a dedicated hue, and are always
stripped from serialized DOCX output until accepted. Accepting a suggestion
converts it into a normal tracked change authored by the accepting user (or, for
a whole inserted table, applies it directly since OOXML has no tracked
representation for it); rejecting inverse-applies it.

Suggested mode covers inline text/format operations (`replaceInBlock`,
`replaceRange`, `formatRange`) and block/table structural operations
(`insertAfterBlock`, `insertBeforeBlock`, `replaceBlock`, `deleteBlock`,
`insertSignatureTable`, `insertTableRow`, `deleteTableRow`, `insertTableColumn`,
`deleteTableColumn`). Whole-node inserts are stripped entirely; suggested
deletes serialize as though they never happened; the strip is the single
`fromProseDoc`/`extractBlocks` boundary every serialization path funnels through.
Cell merge/split and comment operations remain `unsupportedMode`.

New core commands (`getSuggestions`, `acceptSuggestion`, `acceptAllSuggestions`,
`rejectSuggestion`, `rejectAllSuggestions`, `findSuggestionRange`) and
editor-ref methods (`getSuggestions`, `acceptSuggestion` returning
`{ accepted, appliedAs }`, `rejectSuggestion`, `scrollToSuggestion`) expose the
layer to hosts; `getSuggestions` reports each suggestion's kinds and `appliedAs`
(`"tracked"` vs `"direct"`). The React and Vue adapters expose the same ref
surface (the Nuxt module re-exports it).

Tracked changes also gain an optional `initials` field, carried through the
model and the ProseMirror marks/node attrs for UI attribution (hover, accept
authoring). It is intentionally NOT serialized onto `w:ins`/`w:del`/`w:*PrChange`
or table row/cell markers — `w:initials` is not part of ECMA-376
`CT_TrackChange`, so output stays schema-strict — but the parser remains tolerant
of it if an external document supplies one.
