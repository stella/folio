# Sequencing wire fixtures

`batches-v15.json` and `ops/__tests__/__fixtures__/ops-v15.json` pin the
current semantic operation schema. Version 15 adds paragraph insertion at
`beforeTable`/`afterTable` boundaries and exact empty paragraph replacements.
Table-side insertions record block-scoped effects; existing paragraph-relative
insertions keep their positional effect and payload meaning.

Version 14 and older envelopes receive `unsupportedSchema`. No automatic
conversion or live schema upgrade is performed. Historical fixtures remain
unchanged. The DOCX model and Yjs attribute schema are separate and unchanged
by this operation-schema cutover.

Tracked table-side insertion requires a surviving paragraph for its review
inverse and a proved following boundary: a paragraph, a section-free table
receiver with a paragraph terminator, or a section-free pending-review table.
Section, ambiguous-structure, terminal, opaque and standing-range boundaries
that cannot prove reversible review refuse before publication. Direct insertion
and the existing paragraph-relative variants retain their behavior.

Empty replacements need a surviving local insertion boundary and preserve cell
terminators and section endpoints. They refuse an explicitly owned empty
paragraph-formatting field, which ordinary insertion canonicalization could not
restore exactly.

`batches-v14.json` and `ops/__tests__/__fixtures__/ops-v14.json` pin version 14,
which adds complete section-view snapshots to story lifecycle inverses when
section presence or count changes.

`batches-v13.json` and `ops/__tests__/__fixtures__/ops-v13.json` pin
version 13, which records a direct deletion's kept note mark as its
effect, and its envelopes are otherwise those of version 12.

`batches-v12.json` and `ops/__tests__/__fixtures__/ops-v12.json` pin
version 12, which adds `setContentControl`, which sets a block-level
content control's properties by its id; its envelopes are otherwise those of
version 11.

`batches-v11.json` and `ops-v11.json` pin version 11, which records a
sequenced direct format's cut records as its effect; its envelopes are otherwise
those of version 10.

`batches-v9.json` pins accepted sequencing envelopes. It pairs with
`ops/__tests__/__fixtures__/ops-v9.json`, which includes clipboard package
resources alongside the schema-8 semantic table operations. The shipped
`batches-v8.json` and `ops-v8.json` fixtures remain unchanged. Older schema
envelopes receive a structured `unsupportedSchema` refusal.

Table operations decode alone and beside text operations. Their block-scoped
sequenced effects identify affected paragraphs and source-table aliases.
Independent edits commute; competing structural edits of one table receive
`tableRequiresExclusiveEdit`. Decoder and scope tests derive their exercised
operation kinds from the wire fixtures.
