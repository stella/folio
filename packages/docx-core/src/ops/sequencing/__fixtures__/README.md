# Sequencing wire fixtures

`batches-v6.json` pins accepted sequencing envelopes. Schema 4 and 5 envelopes
with currently supported operation shapes normalize to schema 6; unsupported
shapes retain structured refusals.

Table operations are absent from accepted batch fixtures because table transforms
are not implemented. Whole-table, row, column, merge, split, grid and property
operations (including their structural inverses) require exclusive editing and
return `tableRequiresExclusiveEdit`, including in mixed text/table batches.
Their genuine persisted payloads live in `ops/__tests__/__fixtures__/ops-v6.json`;
sequencing tests exercise every declared table kind from those fixtures.
The genuine `ops-v4.json` journal is read back through both the schema 4 and schema
5 readers and classified by supported, exclusive and unsupported families. No
`ops-v5.json` journal was committed; the genuine `batches-v5.json` artifact is read
back separately. Historical artifacts are preserved.
