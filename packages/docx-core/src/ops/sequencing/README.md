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
would not reach the document the replay did. A `replaced` change whose
rebased change reaches a document with the same paragraphs at the same
lengths carries it as `moves`, so positions still follow it. `transformPosition`
moves a position across sequenced batches by paragraph id and offset, so an
editor maps its selection and history without its own coordinates.

`transformBatch` refuses operation pairs whose effects cannot be mapped safely.
An inverse keeps its staleness preconditions: an invalid undo must be dropped,
never applied by removing those preconditions. Supported story addresses derive
from the operation contract.

The unknown-input decoder accepts every operation kind; `BATCH_WIRE_OP_TYPES`
declares them, and a kind added to the schema without a wire form stops the
decoder compiling. Embedded model records are validated within the
decoder's JSON depth and byte bounds: every content node's kind is one the
model defines for its slot, every record carries only the fields its model
type declares, paragraph ids are well formed, scalar formatting values have
their model types and nested content is checked recursively. Deep property
records (borders, shading, drawings, style, numbering-level and section
details) are bounded by the JSON checks alone. Comment records (a creation's
comment, a content edit's paragraphs, the records and anchor paragraphs a
comment inverse restores) are checked the same way; a reply's thread is its
anchor's, never a field of the new comment. A comment's author and date are
the operation's own: the sequencer and every receiver apply them as sent. `validateDocumentBatch`
normalizes optional undefined object fields, and `validateSequencedBatch`
checks assigned revisions and that each effect is the one its operation kind
records.

The position transform rebases text insertions, inline content insertions,
inline record splits and joins, direct deletions, run and paragraph
formatting, paragraph splits and joins, anchored paragraph insertion and
repeated revision decisions. Overlapping revision decisions with opposite
outcomes reject. Text inherited at a paragraph's start takes the run after
it, so a concurrent format from the start covers it in either order.

## Inline insertions

Typing, an inserted atom and the inverse of a deletion that cut no record
compile to `insertContent` with a closed slice (`openStart` and `openEnd` 0) of runs, text and atoms, without zero-width leaves, comment references or
paragraphs (`isInlineInsertion`). Such an insertion moves every later offset
of its paragraph by its width. At its own offset the zero-width gaps decide;
at the same gap the batch sequenced later goes after the one sequenced
earlier, so typing at one point keeps the sequence order whichever peer
applies first. A gap whose count one side leaves absent and the other names
as nonzero cannot be compared and refuses with `unsupportedPair`.

A range that starts at the insertion point moves past the inserted content
and a range that ends there stays before it: the inserted runs carry the
formatting they were typed with, so a concurrent format or deletion never
takes them in, in either order. A range the insertion falls strictly inside
is split around it: a format into two formats, a direct deletion into two
deletions. A tracked deletion around the point, a tracked insertion inside a
direct deletion, and a range with preconditions or joins reaching the point
refuse. Insertion at the start of a paragraph a concurrent join appends
refuses, as for text.

Each sequenced `insertContent` records whether it cut a record carrying an
id (a revision or a content control) in two (`cutIdentified`). Two
insertions at the same gap refuse when the earlier one cut such a record,
since the record's second half would take the id of whichever applied
first; in plain runs they pass.

`splitInline` and `joinInline` only move record boundaries, so offsets map
one to one across them. A concurrent insertion, split or join at the same
offset, a deletion reaching it, a format ending there, a comment anchor
there and a paragraph split or join of the same paragraph refuse, since the
two orders would leave other boundaries.

Typing and other edits join their runs to alike neighbours with seam joins
(`joinInline` with `ifAlike`), which merge only alike runs carrying no ids
and change nothing where the records beside the seam differ or a run
already runs across it. A seam join moves like a point. Content a
concurrent insertion puts at its gap separates the records it would have
joined, whichever applies first, so the join is dropped there: two typings
at one point stay two runs in either order. A concurrent insertion, format
or seam join at its offset passes, and so does a paragraph split or join
away from it. A direct deletion reaching its offset refuses: it may remove
whole records beside the seam, so other records meet there in one order than
in the other. A tracked deletion reaching it removes nothing and passes, as
a pair that records other cuts in the two orders (see below). A split at its
offset, a hard join there and a comment anchor there still refuse.

A slice with open ends merges with the records at its point, and a slice
holding zero-width leaves, comment references or paragraphs depends on
what lies there: such an `insertContent` stays block-scoped.

## Tracked changes beside concurrent cuts

A tracked deletion or insertion records, on its wrapper, how many records it
cut at each edge and how many ran across its whole range, so that rejecting
or accepting it joins them again. Whether a record runs across depends on
what was applied first: when a concurrent edit cuts or merges records at an
edge or inside the range (another tracked edge, an insertion point, a
format's end, an inline split or join, a paragraph split, an inverse
deletion's join), the cut goes to whichever applies first and the two orders
record different joins. A direct deletion that reaches an edge removes what
lay against that edge in one order and moves the edge in the other, and one
covering the change whole leaves the records around it cut in one order
only. Two tracked deletions that overlap (nested, equal or crossing
ranges) leave the shared content to the one sequenced first, under its
author and id: the later one wraps only what is not deleted yet.

These pairs are accepted and rebased like any other pair (`trackedOrderPair`
names their kinds). The authority's order decides the document, and every
peer reaches it by replaying its pending batches over the sequenced ones. A
peer that applied its own edit of such a pair first gets the authority's
document as a `replaced` change, so its undo and redo history is dropped,
and its positions move over the change's `moves`. They are the only pairs
the transform accepts without commuting. Tracked changes apart from every
concurrent cut commute, and so does a direct deletion strictly inside a
tracked one.

## Block-scoped operations

Every other operation is block-scoped: paragraph replacement, inline
replacement and paragraph review, content insertion with open or
document-dependent slices, paragraph deletion, table insertion, deletion and structure edits, numbering
instances, section endpoints and properties, package resources, headers and
footers, story-part restoration, notes and revision metadata. Its footprint is
the paragraphs it addresses, reads or creates and the package parts it
writes; when it is sequenced, its effect records every paragraph it changed,
inserted or removed (`touchedBlocks`), and for a table structure edit every
paragraph of that table. A batch passes a concurrent block-scoped operation
unchanged when it addresses none of those paragraphs, and is refused with
`blockConflict` when it does, with one exception: a paragraph restoration
(`replaceInline`, and each anchor paragraph of `restoreCommentState`) and a
positional edit of that paragraph (typing, an inline insertion, a split or
join of inline records, a deletion or a format within it) meet by the
restoration's span (`restoreTransform.ts`). The span is the stretch between
the longest common prefix and the longest common suffix of the content the
restoration expects and the content it gives, compared leaf by leaf (a
character of run text, or a whole leaf) with the records holding each leaf,
their own fields and where they begin and end: a change of a record's own
fields alone spans that record's leaves, and a change of record boundaries
alone is an empty span at the boundary. An edit reaching into the span is
refused with `blockConflict`. Any other edit passes: rebased over the
restoration, its positions after the span move by the width the span
changes; the restoration rebased over the edit states the paragraph the
edit leaves, before and after it, applying the edit to that paragraph alone,
wherever the paragraph is (a table cell, a header or footer, a note). Both
orders give the same paragraph by construction. A table structure edit is refused with
`tableRequiresExclusiveEdit` only when the concurrent one restructured the
same table; typing in that table's cells passes. Package parts
(`EXCLUSIVE_PARTS`: numbering, styles, relationships, media, sections, headers
and footers, notes, settings, revision metadata and comments) conflict when
one operation writes a part both name, by entry where the operation names one
(a numbering instance, a relationship id, a note, a comment): `partConflict`,
with the part. A comment operation that may add or remove the comment part
relationships writes the whole relationships part. A revision decision reads
the revisions it resolves, and a paragraph replacement, inline replacement,
paragraph review or container replacement writes every revision its new
content holds: an undo that puts a tracked change back conflicts with a
concurrent decision of it, which would otherwise resolve the change in one
order and find it already resolved in the other. Pairs of position-rebased
operations the transform cannot map safely reject with `unsupportedPair`.

## Comments

Comment threads are sequenced in the journal like any other edit. A range or
point comment is rebased by its anchor: its ends move like text over
concurrent insertions, direct deletions, splits and joins, and stay over
formatting, paragraph properties and paragraph insertion. Its comment
reference occupies one offset unit at the range's end (or at the point), so a
concurrent position after it moves past it; typing and other inline
insertions in the anchor's paragraph move the same way. A position at an
anchor's own offset has no affinity to follow and refuses with
`unsupportedPair`, as does
a concurrent deletion or format whose range reaches an anchor. A deletion
that removes a range's text is therefore refused rather than collapsing the
range: the apply contract makes a collapsed range a point comment, which the
other order (a range comment whose text is then deleted) would not give. A
reply, a deletion and a restoration find their anchors in the document they
apply to; their effect records the paragraphs they changed, so a concurrent
edit addressing one of those paragraphs refuses with `blockConflict` (a
restoration meets a positional edit by its span instead, as above), and
the comments they added or removed, replies of a deleted thread among them.

Concurrent comment operations decide by the comments they use. Different
comments pass: a new comment takes its place in the comment list by id, so
concurrent creations land in the same order whichever applies first. On one
comment, a content edit and a resolution pass, a reply passes an edit or a
resolution of its parent, and the same resolution twice is one resolution.
Two replies to one parent pass: a reply's markers follow the parent's (or
its nearest anchored ancestor's) and the markers of the thread's comments
that come before it, in depth-first order with siblings by id, so
concurrent replies take the same places whichever applies first.
Everything else on one comment is refused with `partConflict` naming the
`comments` part: two content edits, opposite resolutions, a deletion against
any other use of its thread (a reply to a deleted parent included), and a
restoration against anything but a content edit of another comment (it puts
records and part relationships back as captured). A reply or a resolution
after a concurrent deletion that removed a comment part relationship is
refused naming `relationships`, since it would add the relationship under
another id than it does before the deletion. Comment ids come from the
actor's lease like every other id.

Undo inverses are operations of the same kinds and follow the same rules.
Rich inverse preconditions stay intact and may refuse an undo after another
edit.

Partially overlapping writes to the same run property require additional
structural effects and currently reject. Nonzero marker-gap changes and
identified ranges crossing a split also reject. Inserting within tracked-deleted content
is refused by the apply contract. A join can preserve two adjacent runs; an
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
an authority entry is deleted or two claim one revision. Every Yjs client
picks its own client id, so this check catches a client that writes the
journal by mistake, not one that writes it under the authority's id on
purpose: the room trusts its participants, and a host that does not trust
one drops that client's updates touching the journal, the answers or the
room map before it relays them. A lease call that the room can no longer
forward, because the transport closed or stopped following the journal, is
answered with an `unavailable` refusal. The host persists
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
positions. Deliveries may reorder a sender's states; a tab room keeps each
sender's latest by its clock, so a peer that starts following presence later
is handed the latest state, not the one that arrived last.
