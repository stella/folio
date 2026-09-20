---
"@stll/folio-core": minor
---

Carry a section break on one attr.

Two paragraph attrs carried one state: `_sectionProperties`, the parsed record the save leg preferred, and `sectionBreakType`, which the editor commands wrote. `removeSectionBreak` cleared the type and left the record, so a parsed break it claimed to remove was still saved; `insertSectionBreak` could not retype a parsed break without the save leg minting a fresh record and dropping that section's page size, margins, columns and header references; and a paragraph holding the type alone had no record for a split to be read against, so both halves minted their own and the document gained a section.

`_sectionProperties` is now the whole state. The break type is a field of the record (`w:type`, ECMA-376 Part 1 §17.6.22), derived through `sectionBreakTypeOf` for the layout bridge, compare, the change tracker, the DOM and the toolbar commands, and `ParagraphAttrs.sectionBreakType` is gone from the schema. A break the editor inserts mints one record and shares it by reference, exactly as a parsed one is shared, so the save leg's rule (among the paragraphs holding one record, the last in document order writes it) covers both.

Backspace at the start of the paragraph *after* a break now deletes the break, as Word does: §17.6.18 puts the section's properties on the mark the join consumes, and the paragraphs it governed fall to the following section, whose `w:sectPr` governs them from then on.

Collaboration snapshots carry an attr-schema version, bumped to 9. An older snapshot that states `sectionBreakType` and no record has the record minted for it on load; without the step ProseMirror would drop the attr its schema no longer declares and the save would write one `w:sectPr` fewer.
