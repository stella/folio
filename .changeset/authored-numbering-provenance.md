---
"@stll/folio-core": minor
"@stll/folio-agents": minor
"@stll/folio-cli": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Breaking: preserve stated numbering references separately from style inheritance. Collaboration attr schema 13 refuses older snapshots with ambiguous numbering ownership; rebuild those snapshots from the saved DOCX.

Numbering requests now use named inherit, none, levelOnly, reference, or newList variants. Replace null cancellation with { kind: "none" }, and listLevel patches with levelOnly numbering. Read statedNumbering for authored state and listReference for effective membership; listLevel readback is removed.

Paragraph blocks require statedNumbering; diagnostic blocks expose no paragraph numbering or formatting. Narrow on block.kind before reading those fields.

Redline insertions import their referenced style closure and numbering through collision-safe resource owners while preserving stated inheritance. An unimportable resource closure raises GenerateRedlineDocxResourceImportError before any body operation, preserving the base document.

Source-undefined style references stay verbatim when also undefined in the base. Styles that would bind to an unrelated base definition are cleared and reported in referenceWarnings with their revised story position. Dangling numbering remains normalized to none by the existing parser, so inserted paragraphs stay unnumbered through save/reopen.

Breaking: undefinedStyles is replaced by the required undefinedReferences option, with no default or alias. Every execution caller must pass { undefinedReferences: "refuse" } for editing or { undefinedReferences: "keep" } for comparison/redline. Both reviewer methods and React/Vue operation refs require an explicit policy, applying it to style and numbering references alike.
