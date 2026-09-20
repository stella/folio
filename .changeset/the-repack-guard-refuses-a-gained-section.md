---
"@stll/folio-core": patch
---

Refuse a save that adds a section the editor never authored.

The repack fidelity guard read one direction of the section count: it refused a save that dropped a section and said nothing about one that added a section, which is why a split paragraph's duplicated `w:sectPr` reached the file without anything noticing.

A gain is legitimate, since the editor inserts breaks, so the guard asks what the model holds rather than what the original held. Two paragraphs over the same `SectionProperties` record are one section's split halves, never two sections: that fails with `DocxDuplicateSectionCarrierError`, which names the paragraph. The package must also state exactly as many `w:sectPr` elements as the model has records, because a record the serializer fails closed on would otherwise pass as a smaller gain rather than a loss.
