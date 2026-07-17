---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-agents": minor
---

Add a first-class suggestion layer to tracked changes. AI-proposed edits can be
applied with the new `"suggested"` apply mode: they render with the
tracked-change grammar but a dotted stroke and a dedicated hue, and are always
stripped from serialized DOCX output until accepted. Accepting a suggestion
converts its marks into normal tracked changes authored by the accepting user;
rejecting inverse-applies them. New core commands (`getSuggestions`,
`acceptSuggestion`, `acceptAllSuggestions`, `rejectSuggestion`,
`rejectAllSuggestions`, `findSuggestionRange`) and editor-ref methods
(`getSuggestions`, `acceptSuggestion`, `rejectSuggestion`, `scrollToSuggestion`)
expose the layer to hosts. Suggested mode currently covers inline text and
formatting operations (`replaceInBlock`, `replaceRange`, `formatRange`);
block- and table-level operations report `unsupportedMode`.
