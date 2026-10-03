# Sequencing wire fixtures

`batches-v8.json` pins accepted sequencing envelopes. Older schema envelopes
receive a structured `unsupportedSchema` refusal; historical artifacts remain
unchanged.

Table operations are absent from accepted batch fixtures because table transforms
are not implemented. Whole-table, row, column, merge, split, grid and property
operations (including their structural inverses) require exclusive editing and
return `tableRequiresExclusiveEdit`, including in mixed text/table batches.
Their persisted payloads live in `ops/__tests__/__fixtures__/ops-v8.json`;
sequencing tests exercise every declared table kind from those fixtures.
