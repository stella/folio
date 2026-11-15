# Sequencing wire fixtures

`batches-v12.json` and `ops/__tests__/__fixtures__/ops-v12.json` pin the
current schema: version 12 adds `setContentControl`, which sets a block-level
content control's properties by its id; its envelopes are otherwise those of
version 11.

`batches-v11.json` and `ops-v11.json` pin version 11, which records a
sequenced direct format's cut run as its effect; its envelopes are otherwise
those of version 10.

`batches-v9.json` pins accepted sequencing envelopes. It pairs with
`ops/__tests__/__fixtures__/ops-v9.json`, which includes clipboard package
resources alongside the schema-8 semantic table operations. The shipped
`batches-v8.json` and `ops-v8.json` fixtures remain unchanged. Older schema
envelopes receive a structured `unsupportedSchema` refusal.

Table operations are absent from accepted batch fixtures because table transforms
are not implemented. Whole-table, row, column, merge, split, grid and property
operations (including their structural inverses) require exclusive editing and
return `tableRequiresExclusiveEdit`, including in mixed text/table batches.
Their persisted payloads live in `ops/__tests__/__fixtures__/ops-v9.json`;
sequencing tests exercise every declared table kind from those fixtures.
