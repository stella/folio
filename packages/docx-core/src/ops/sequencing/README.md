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

Each authoritative answer returns how the client's document moved: a `rebased`
change, transformable like a sequenced batch, or a `replaced` change, exact
for the document only, when a pending batch was dropped or the rebased change
would not reach the document the replay did. `transformPosition`
moves a position across sequenced batches by paragraph id and offset, so an
editor maps its selection and history without its own coordinates.

`transformBatch` refuses operation pairs whose effects cannot be mapped safely.
An inverse keeps its staleness preconditions: an invalid undo must be dropped,
never applied by removing those preconditions. Supported story addresses derive
from the operation contract; this schema has no `addComment` primitive.

The unknown-input decoder accepts every operation kind except comment
operations (`UNSENT_OP_TYPES`); `BATCH_WIRE_OP_TYPES` declares the accepted
kinds, and a kind added to the schema without a wire form or an unsent entry
stops the decoder compiling. Embedded model records are validated within the
decoder's JSON depth and byte bounds: every content node's kind is one the
model defines for its slot, every record carries only the fields its model
type declares, paragraph ids are well formed, scalar formatting values have
their model types and nested content is checked recursively. Deep property
records (borders, shading, drawings, style, numbering-level and section
details) are bounded by the JSON checks alone. `validateDocumentBatch`
normalizes optional undefined object fields, and `validateSequencedBatch`
checks assigned revisions and that each effect is the one its operation kind
records.

The position transform rebases text insertions, direct deletions, run and
paragraph formatting, paragraph splits and joins, anchored paragraph
insertion and repeated revision decisions. Overlapping revision decisions with
opposite outcomes reject. Text inherited at a paragraph's start takes the run
after it, so a concurrent format from the start covers it in either order.

Every other operation is block-scoped: paragraph replacement, inline
replacement and paragraph review, inline splits and joins, content insertion,
paragraph deletion, table insertion, deletion and structure edits, numbering
instances, section endpoints and properties, package resources, headers and
footers, story-part restoration, notes and revision metadata. Its footprint is
the paragraphs it addresses, reads or creates and the package parts it
writes; when it is sequenced, its effect records every paragraph it changed,
inserted or removed (`touchedBlocks`), and for a table structure edit every
paragraph of that table. A batch passes a concurrent block-scoped operation
unchanged when it addresses none of those paragraphs, and is refused with
`blockConflict` when it does. A table structure edit is refused with
`tableRequiresExclusiveEdit` only when the concurrent one restructured the
same table; typing in that table's cells passes. Package parts
(`EXCLUSIVE_PARTS`: numbering, styles, relationships, media, sections, headers
and footers, notes, settings and revision metadata) conflict when one
operation writes a part both name, by entry where the operation names one
(a numbering instance, a relationship id, a note): `partConflict`, with the
part. Pairs of position-rebased operations the transform cannot map safely
reject with `unsupportedPair`.

Undo inverses are operations of the same kinds and follow the same rules.
Rich inverse preconditions stay intact and may refuse an undo after another
edit.

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

## Joining a canonical session to a room

`@stll/folio-core` binds one canonical editing session to a room
(`joinCanonicalRoom`, controller/canonicalCollaboration.ts) over a transport
that carries batches and leases only (`CanonicalRoomTransport`): the room's
base snapshot identity (the SHA-256 of the DOCX it was seeded from, which
revision 0 is), its head revision, `submit` answering an acknowledgement or
a rejection, a subscription delivering the journal after a revision, and the
actor's id lease. `createCanonicalMemoryRoom` is the in-memory reference
transport around `createSequencer`.

A session joins only the room seeded from the DOCX it opened, refusing any
other snapshot, and only before it is edited; it creates its session with the
granted lease and replays the journal from revision 0. Each commit the
session publishes (its committed-batch feed, `onCommitted`) is enqueued on the
optimistic client under the commit's operation id, at the client's head
revision, and submitted one at a time. A sequenced batch goes to the client
first, which rebases the pending batches over it, and the change the client
reports enters the session through `prepareRemote`. Deliveries may repeat or
arrive out of order: revisions already applied are ignored and later ones
wait for the gap. A rejected batch, or a pending batch the client dropped, is
taken back the same way and reported to the host (`localEditRefused`), never
lost silently. After a disconnection the unanswered batch is submitted again
under the same operation id, so the sequencer answers it once. The controller
asks for another lease block when a kind runs low. A remote batch waits while
a composition is pending. The controller holds no Yjs state.

## Room transports

A room has exactly one authority: one sequencer over the room's base
document. Its durable state is an ordered list of records, every sequenced
batch and every leased id block (`CanonicalRoomRecord`).
`restoreCanonicalRoomAuthority` rebuilds the sequencer from them: the leased
blocks go back into the lease table, then every batch is submitted again at
its own revision and must sequence there. An authority stores the records a
call added before it broadcasts or answers anything, so a successor never
forgets an edit somebody heard of, and a batch submitted again under an
operation id it already sequenced gets its first answer.

`openCanonicalTabRoom` (controller/canonicalTabRoom.ts) shares a room
between the tabs of one origin. The authority runs in the tab that holds the
room's Web Lock (`navigator.locks`), a leader tab rather than a SharedWorker:
the library ships no worker script a host would have to serve, it works
where SharedWorker does not, and a durable journal is needed either way, so
the room outlives the tabs. Every tab, the leader included, reaches the
authority over a BroadcastChannel. When the leading tab goes away the lock
passes to a waiting tab, which reads the durable journal
(`openCanonicalIndexedDbJournal`, one IndexedDB database per origin),
rebuilds the sequencer and stores a higher leadership term before it
announces itself. Clients follow only a higher term: the old leader's
unanswered submissions count as lost to a disconnection, and the controller
submits them again under the same operation id.

`openCanonicalYjsRoom` (controller/canonicalYjsRoom.ts) carries a room in a
Yjs document, for hosts that already sync one. The document is a transport,
not the shared model: a client appends submissions and lease requests to an
inbox array, and the authority helper a host server runs on its copy of the
document (`serveCanonicalYjsRoom`, for example in a y-websocket or Hocuspocus
server) takes each request in order and, in one transaction, appends the
records to a journal array, writes the answer to an answers map keyed by
actor and operation id, and removes the request. Only the authority writes
the journal, under its own Yjs client id (`CANONICAL_YJS_AUTHORITY` unless
the host picks another): a client refuses an entry any other Yjs client
wrote (`foreignJournalEntry`) and goes on, and stops following the room if
an authority entry is deleted or two claim one revision. The host persists
the document's updates before it sends them on, so a restarted authority
finds every entry a client saw. The module imports no Yjs code; it works on
the host's own document and awareness. A serverless room (y-webrtc) has no
authority and cannot carry a canonical room until it gets one: a peer that
is always present or a server that joins the room.

## Presence

With the `presence` option, `joinCanonicalRoom` shares the peer's selection
over the transport's presence channel (the tab room's BroadcastChannel, the
Yjs room's awareness): canonical positions (story, paragraph id, offset), the
revision the peer had applied, and its user's name and color. A viewer moves
every other peer's selection over each sequenced batch it applies after that
revision, other than the peer's own, with the transform the session moves
its own selection with, and `presence.peers()` gives each peer's selection
in the viewer's projection (`positionAt`). A peer publishes again after each
remote edit and each revision the room confirms, and the host calls
`presence.update()` when the selection moves. Presence uses no Yjs relative
positions.
