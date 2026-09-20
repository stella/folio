---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep the attributes a modelled property element carries that the model has no field for.

A property element is an attribute bag, and the child dispatcher decides one whole: a handler either reads the element or hands back its bytes. `<w:ind w:leftChars="100"/>` survived because the reader took nothing from it, while `<w:ind w:left="720" w:leftChars="100"/>` — what a document actually carries — was modelled and lost the character unit. The same went for `w:spacing`'s line counts, `w:framePr`'s `w:hRule` and `w:anchorLock`, and the three attributes describing a `w:shd` pattern colour.

The attribute remainder now rides the record that holds the element's modelled fields: `ShadingProperties`, `BorderSpec`, `TabStop` and `ParagraphFormatting.frame` gain `preservedAttributes`, and `w:ind` and `w:spacing`, which the model flattened into `ParagraphFormatting`, gain one remainder each there. The predicate is derived from the model rather than written beside each reader — `propertyElementAttributes.ts` holds one `as const satisfies` table per record — so a field added without an attribute to name does not compile.

The container-survival census measured one attribute at a time and so could not see the defect at all. It now states each attribute pair a second time beside one the element models, and the pair survives only when it survives both.
