# Operation batches

The `@stll/docx-core/ops` entry point provides versioned batches, a shared
position transform, and deterministic reference implementations of a sequencer
and an optimistic client. They perform no I/O; the caller owns authentication,
storage, delivery and notifications.

A batch carries `schema`, `opId`, `actor`, `baseRev` and ordered `ops`. The
sequencer assigns its `revision`. `actor` is an opaque identity supplied by the
caller. Offsets use the operation model's logical UTF-16 space: fields and
inline atoms occupy one unit, deleted tracked content still occupies space,
and `zeroWidthBefore` distinguishes gaps between markers.

The sequencer returns the original acknowledgement or rejection for every
resubmission of an `opId`. An accepted batch applies atomically and appears
in the broadcast journal. A rejection leaves the document and journal unchanged.
Sequenced effects record document-dependent split identities and join lengths,
so replay does not infer positions from a later document.

`transformBatch` refuses operation pairs whose effects cannot be mapped safely.
An inverse keeps its staleness preconditions: an invalid undo must be dropped,
never applied by removing those preconditions. Supported story addresses derive
from the operation contract; this schema has no `addComment` primitive.

The unknown-input decoder accepts a bounded subset of model records: scalar
formatting, language properties, plain runs and basic inline atoms. It refuses
unsupported embedded records and operation kinds. `BATCH_WIRE_OP_TYPES` declares
the accepted kinds; `validateDocumentBatch` normalizes optional undefined object
fields, and `validateSequencedBatch` checks assigned revisions and effects.

The initial transform supports text insertions, direct deletions, formatting,
paragraph splits and joins, anchored paragraph insertion, and repeated revision
decisions. Numbering and table structure require exclusive edits. Overlapping
revision decisions with opposite outcomes reject. Unspecified pairs reject.
Rich inverse preconditions stay intact and may refuse an undo after another edit.

Partially overlapping writes to the same run property require additional
structural effects and currently reject. Nonzero marker-gap changes and
identified ranges crossing a split also reject. Inserting within tracked-deleted content
is refused by the apply contract. A join can preserve two adjacent runs; text
insertion at that boundary requires an affinity the current text operation
does not express, so the transform refuses ambiguous boundary insertion.
