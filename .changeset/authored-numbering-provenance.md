---
"@stll/folio-core": minor
"@stll/folio-agents": minor
"@stll/folio-cli": minor
---

Breaking: preserve stated numbering references separately from style inheritance. Collaboration attr schema 13 refuses older snapshots with ambiguous numbering ownership; rebuild those snapshots from the saved DOCX.

Numbering requests now use named inherit, none, levelOnly, reference, or newList variants. Replace null cancellation with { kind: "none" }, and listLevel patches with levelOnly numbering. Read statedNumbering for authored state and listReference for effective membership; listLevel readback is removed.

Paragraph blocks require statedNumbering; diagnostic blocks expose no paragraph numbering or formatting. Narrow on block.kind before reading those fields.

Redline insertions import their referenced style closure and numbering through collision-safe resource owners while preserving stated inheritance. An unimportable resource closure raises GenerateRedlineDocxResourceImportError before any body operation, preserving the base document.

Source-undefined style and numbering references stay verbatim when also undefined in the base. References that would bind to an unrelated base definition are cleared and reported in referenceWarnings with their revised story position. The execution option undefinedStyles is replaced by undefinedReferences, applying the same keep/refuse policy to both resource kinds.
