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
resubmission of an `opId` by the same `actor`; another actor reusing it is
rejected as a conflict. The optimistic client likewise refuses a broadcast that
reuses a pending `opId` under another actor. An accepted batch applies atomically and appears
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

## Id leases

Concurrent actors must never mint the same paragraph (`w14:paraId`), revision
(`w:id`), content-control or comment id. The sequencer leases each actor
disjoint blocks of each kind (`grantIdLease`): by default 4096 paragraph ids,
4096 revision ids, 1024 content-control ids and 1024 comment ids. Revision,
content-control and comment ids share one numeric space, so their blocks never
overlap either. A granted block holds no id the room has held. An actor mints
upward through its lease and never mints an id twice; without a lease, the
allocators keep minting the lowest free ids.

The sequencer rejects a batch that introduces an id outside its actor's lease
(`outsideIdLease`, with the id kind). An id the room held before is accepted
from any actor, as is an id at or below what its lessee already minted: both
are restorations, such as an undo, that nobody can mint again.

`grantIdLease` is idempotent: it returns the actor's remaining blocks, so a
reloading actor gets its lease back. `renewIdLease` adds one block of a kind
while the actor holds at most one block of unused ids of it, and refuses with
`notDue` otherwise. The lease table (`idLeases`) is part of the sequenced
state: it is stored with the journal and passed back to `createSequencer` on a
restart. A restart from a compacted document forgets ids that were removed
before it; restoring such an id is refused rather than risking a collision.

Paragraph ids stay below 0x80000000 and `w:id` within a signed 32-bit integer.
When no free block is left below that limit, granting or renewing refuses with
`exhausted`; ids never wrap. An actor whose lease runs out cannot mint: its
edits that need new ids are refused until the room is compacted into a new
base document.
