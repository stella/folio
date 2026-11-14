import { Result } from "better-result";
/**
 * Peers editing one room concurrently with every kind of canonical edit,
 * comment threads among them, through the real sequencer and optimistic
 * client, with every batch crossing the JSON wire: all of them end on the
 * authority's bytes, comments included, no edit is lost without a notice,
 * and every concurrent batch the authority accepted commutes with the
 * journal tail it was rebased over.
 */
import { describe, expect, mock, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../../../test/property-coverage";
import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import type { Document } from "../../../model/document";
import { serializeDocumentToDocx } from "../../../serialize/docx";
import { createEditorIntentIdAllocator } from "../../editorIntent";
import { equalForStaleness } from "../../equality";
import {
  ID_LEASE_KINDS,
  introducedIds,
  leasedIdCensus,
  trimIdLease,
} from "../../idLease";
import {
  type DeleteRangeOp,
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertContentOp,
  type TextPosition,
} from "../../types";
import { restoreDocumentOp } from "../../wire";
import * as applyBatchModule from "../applyBatch";
import { applyBatch } from "../applyBatch";
import { type ClientDocumentChange, createClient } from "../client";
import * as commentTransform from "../commentTransform";
import {
  anchorPoints,
  isAnchoredComment,
  isCommentOp,
} from "../commentTransform";
import * as envelope from "../envelope";
import {
  BATCH_REJECTION_REASONS,
  type BatchRejectionReason,
  type DocumentBatch,
  type SequencedBatch,
  validateDocumentBatch,
} from "../envelope";
import * as heldRevisions from "../heldRevisions";
import * as inlineTransform from "../inlineTransform";
import { isInlineInsertion, sameParagraph } from "../points";
import { isBlockScoped } from "../scope";
import { type BatchSubmission, createSequencer } from "../sequencer";
import * as trackedTransform from "../trackedTransform";
import { transformBatch } from "../transform";
import {
  canonicalBase,
  compileCanonicalEdit,
  type EditAuthor,
  type EditKind,
  type EditSeed,
  editSeedArbitrary,
} from "./canonicalEdits";

const NUM_RUNS = 25;

type Step =
  | { type: "edit"; peer: number; seed: EditSeed }
  | { type: "undo"; peer: number }
  | { type: "submit"; peer: number }
  | { type: "deliver"; seed: number }
  /**
   * The peer's connection drops before its answer arrives: the answer is
   * lost and the batch in flight is submitted again under its operation id.
   */
  | { type: "reconnect"; peer: number };

type Peer = {
  author: EditAuthor;
  client: ReturnType<typeof createClient>;
  /** The inverse of this peer's last edit, as applying it returned. */
  undo: readonly DocumentOp[] | undefined;
};

type Delivery =
  | { peer: number; type: "result"; result: BatchSubmission }
  | { peer: number; type: "broadcast"; batch: SequencedBatch };

const wire = <Value>(value: Value): Value => JSON.parse(JSON.stringify(value));

/** Refusals that would mean the wire, the base or the leases failed. */
const FORBIDDEN: ReadonlySet<BatchRejectionReason> = new Set([
  BATCH_REJECTION_REASONS.INVALID_BATCH,
  BATCH_REJECTION_REASONS.UNSUPPORTED_SCHEMA,
  BATCH_REJECTION_REASONS.STALE_BASE,
  BATCH_REJECTION_REASONS.OUTSIDE_ID_LEASE,
]);

type Outcome = {
  enqueued: string[];
  sequenced: string[];
  rejected: BatchRejectionReason[];
  dropped: string[];
  documents: Document[];
  authority: Document;
  /** Accepted batches rebased over a tail, and how many were proven to commute. */
  rebased: number;
  proven: number;
  scopedRebased: number;
  /** Accepted rebased batches with a comment operation in them or their tail. */
  commentRebased: number;
  /** Accepted rebased batches whose result differs from the commuted order. */
  diverged: string[];
  /** Accepted rebased batches with a tracked deletion in one paragraph as one in their tail. */
  trackedPairs: number;
  /** Concurrent pairs in one paragraph or thread, accepted and refused. */
  pairs: PairCounts;
  /** Why two typings refused each other: the pair's types and message. */
  typingRefusals: string[];
  /** Typing pairs refused because one side cuts or rejoins an identified record. */
  identifiedTypingRefusals: string[];
  /** Accepted batches that did not commute, all of a declared tracked pair kind. */
  trackedOrder: Record<trackedTransform.TrackedOrderPair, number>;
  /** Refusals of a declared tracked pair kind: the bar is none. */
  trackedRefusals: string[];
  /** Changes a peer with pending batches got as a replacement, by the pair kinds they met. */
  replaced: Tally;
  /** Changes a peer with pending batches got rebased, by the pair kinds they met. */
  rebasedChanges: Tally;
  /** Batches submitted again after a lost answer. */
  resubmitted: number;
};

const noTrackedOrder = (): Record<
  trackedTransform.TrackedOrderPair,
  number
> => ({
  cutAtTrackedChange: 0,
  deletionAtTrackedEdge: 0,
  overlappingDeletions: 0,
});
const TRACKED_ORDER_KINDS = Object.values(trackedTransform.TRACKED_ORDER_PAIRS);

/**
 * The declared tracked pairs a batch forms with the tail it was rebased
 * over, each compared in one coordinate space: the batch moved over the
 * tail batches sequenced before the one it meets.
 */
const declaredOver = (
  batch: DocumentBatch,
  tail: readonly SequencedBatch[],
): Set<trackedTransform.TrackedOrderPair> => {
  const kinds = new Set<trackedTransform.TrackedOrderPair>();
  for (const [index, remote] of tail.entries()) {
    const moved = transformBatch(batch, tail.slice(0, index));
    if (moved.isErr()) break;
    for (const kind of trackedOrderPairs(moved.value.ops, remote.ops))
      kinds.add(kind);
  }
  return kinds;
};

// The batch's operations where they meet the tail batch that refuses them.
const opsAtRefusal = (
  batch: DocumentBatch,
  tail: readonly SequencedBatch[],
): readonly DocumentOp[] => {
  const index = tail.findIndex((_, at) =>
    transformBatch(batch, tail.slice(0, at + 1)).isErr(),
  );
  if (index <= 0) return batch.ops;
  const moved = transformBatch(batch, tail.slice(0, index));
  return moved.isOk() ? moved.value.ops : batch.ops;
};

// Where two values first differ, for a failure message.
const differenceAt = (
  left: unknown,
  right: unknown,
  path: string,
): string | undefined => {
  if (equalForStaleness(left, right)) return undefined;
  if (
    typeof left === "object" &&
    left !== null &&
    typeof right === "object" &&
    right !== null
  )
    for (const key of new Set([...Object.keys(left), ...Object.keys(right)])) {
      const found = differenceAt(
        Reflect.get(left, key),
        Reflect.get(right, key),
        `${path}.${key}`,
      );
      if (found !== undefined) return found;
    }
  return `${path}: ${JSON.stringify(left)?.slice(0, 300) ?? "undefined"} != ${JSON.stringify(right)?.slice(0, 300) ?? "undefined"}`;
};

// An insertion and a deletion, either way round.
const insertionAgainstDeletion = (
  op: DocumentOp,
  other: DocumentOp,
): boolean => {
  const inserts = (candidate: DocumentOp) =>
    candidate.type === DOCUMENT_OP_TYPES.INSERT_CONTENT ||
    candidate.type === DOCUMENT_OP_TYPES.INSERT_TEXT;
  return (
    (inserts(op) && other.type === DOCUMENT_OP_TYPES.DELETE_RANGE) ||
    (inserts(other) && op.type === DOCUMENT_OP_TYPES.DELETE_RANGE)
  );
};

// A split or a join that is not a seam join: it cuts or merges records exactly.
const cutsRecordsExactly = (op: DocumentOp): boolean =>
  op.type === DOCUMENT_OP_TYPES.SPLIT_INLINE ||
  (op.type === DOCUMENT_OP_TYPES.JOIN_INLINE && op.ifAlike !== true);

// An inverse whose preconditions or joins pin the range it restores.
const holdsPreconditions = (op: DocumentOp): boolean => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
      return op.expected !== undefined || op.join !== undefined;
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return (
        op.expected !== undefined ||
        op.joinStart !== undefined ||
        op.joinEnd !== undefined
      );
    default:
      return false;
  }
};

/** Changes a peer with pending batches got, by the pair kinds they met. */
type Tally = Record<PairKind | "other", number>;
const noTally = (): Tally => ({
  typingTyping: 0,
  typingFormat: 0,
  typingDeletion: 0,
  typingComment: 0,
  replyReply: 0,
  suggestedDeletions: 0,
  suggestedEdits: 0,
  suggestionAgainstDirect: 0,
  other: 0,
});

// The declared order-dependent tracked pairs between a batch's operations
// and some concurrent ones.
const trackedOrderPairs = (
  ops: readonly DocumentOp[],
  others: readonly DocumentOp[],
): Set<trackedTransform.TrackedOrderPair> => {
  const kinds = new Set<trackedTransform.TrackedOrderPair>();
  for (const op of restoredOps(ops))
    for (const other of restoredOps(others)) {
      const kind = trackedTransform.trackedOrderPair(op, other);
      if (kind !== undefined) kinds.add(kind);
    }
  return kinds;
};

/**
 * Concurrent pairs in one paragraph (or on one thread) a batch met in the
 * tail it was rebased over: inline insertions against each other, a format,
 * a deletion and a comment anchor; two replies to one parent.
 */
const TYPING_PAIR_KINDS = [
  "typingTyping",
  "typingFormat",
  "typingDeletion",
  "typingComment",
  "replyReply",
] as const;
/**
 * Suggesting pairs in one paragraph: two tracked deletions, two other
 * tracked changes (two suggesting typists among them), and a tracked change
 * against a direct edit.
 */
const SUGGESTING_PAIR_KINDS = [
  "suggestedDeletions",
  "suggestedEdits",
  "suggestionAgainstDirect",
] as const;
const PAIR_KINDS = [...TYPING_PAIR_KINDS, ...SUGGESTING_PAIR_KINDS] as const;
type PairKind = (typeof PAIR_KINDS)[number];
type PairCounts = Record<PairKind, { accepted: number; refused: number }>;
const noPairs = (): PairCounts => ({
  typingTyping: { accepted: 0, refused: 0 },
  typingFormat: { accepted: 0, refused: 0 },
  typingDeletion: { accepted: 0, refused: 0 },
  typingComment: { accepted: 0, refused: 0 },
  replyReply: { accepted: 0, refused: 0 },
  suggestedDeletions: { accepted: 0, refused: 0 },
  suggestedEdits: { accepted: 0, refused: 0 },
  suggestionAgainstDirect: { accepted: 0, refused: 0 },
});

// The paragraph position an operation of a counted pair addresses.
const pairPoint = (op: DocumentOp): TextPosition | undefined => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
      return op.at;
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return op.from;
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
      return isAnchoredComment(op) ? anchorPoints(op).at(0) : undefined;
    default:
      return undefined;
  }
};

// Direct typing or an inserted atom: an inline insertion recording no change.
const isTyping = (op: DocumentOp): op is InsertContentOp =>
  isInlineInsertion(op) && op.revision === undefined;

const typingAgainst = (
  typed: DocumentOp,
  other: DocumentOp,
): PairKind | undefined => {
  if (!isTyping(typed)) return undefined;
  const point = pairPoint(other);
  if (point === undefined || !sameParagraph(point, typed.at)) return undefined;
  switch (other.type) {
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
      return isTyping(other) ? "typingTyping" : undefined;
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return "typingFormat";
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
      return "typingDeletion";
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
      return "typingComment";
    default:
      return undefined;
  }
};

// Whether an operation records a text change as tracked.
const tracksText = (op: DocumentOp): boolean =>
  (op.type === DOCUMENT_OP_TYPES.DELETE_RANGE ||
    op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT ||
    op.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS) &&
  op.revision !== undefined;

const suggestingPair = (
  left: DocumentOp,
  right: DocumentOp,
): PairKind | undefined => {
  const leftPoint = pairPoint(left);
  const rightPoint = pairPoint(right);
  if (
    leftPoint === undefined ||
    rightPoint === undefined ||
    left.type === DOCUMENT_OP_TYPES.CREATE_COMMENT ||
    right.type === DOCUMENT_OP_TYPES.CREATE_COMMENT ||
    !sameParagraph(leftPoint, rightPoint)
  )
    return undefined;
  const tracked = [tracksText(left), tracksText(right)];
  if (tracked.every(Boolean))
    return left.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
      right.type === DOCUMENT_OP_TYPES.DELETE_RANGE
      ? "suggestedDeletions"
      : "suggestedEdits";
  return tracked.some(Boolean) ? "suggestionAgainstDirect" : undefined;
};

const pairKindsOf = (left: DocumentOp, right: DocumentOp): PairKind[] => {
  if (
    left.type === DOCUMENT_OP_TYPES.CREATE_COMMENT &&
    right.type === DOCUMENT_OP_TYPES.CREATE_COMMENT &&
    left.anchor.kind === "reply" &&
    right.anchor.kind === "reply" &&
    left.anchor.parentId === right.anchor.parentId
  )
    return ["replyReply"];
  return [
    typingAgainst(left, right) ?? typingAgainst(right, left),
    suggestingPair(left, right),
  ].filter((kind): kind is PairKind => kind !== undefined);
};

const restoredOps = (ops: readonly DocumentOp[]): DocumentOp[] =>
  ops.flatMap((op) => {
    const restored = restoreDocumentOp(op);
    return restored.isOk() ? [restored.value] : [];
  });

// The kinds of pair between a batch's operations and some concurrent ones.
const pairKinds = (
  ops: readonly DocumentOp[],
  others: readonly DocumentOp[],
): Set<PairKind> => {
  const kinds = new Set<PairKind>();
  for (const op of restoredOps(ops))
    for (const other of restoredOps(others)) {
      for (const kind of pairKindsOf(op, other)) kinds.add(kind);
    }
  return kinds;
};

// Operation kinds with the fields that decide a typing pair.
const describeOps = (ops: readonly DocumentOp[]): string =>
  ops
    .map((op) => {
      switch (op.type) {
        case DOCUMENT_OP_TYPES.JOIN_INLINE:
          return `joinInline@${String(op.at.offset)}/${String(op.depth)}${op.ifAlike === true ? "/seam" : ""}`;
        case DOCUMENT_OP_TYPES.SPLIT_INLINE:
          return `splitInline@${String(op.at.offset)}/${String(op.depth)}`;
        case DOCUMENT_OP_TYPES.INSERT_CONTENT:
          return `insertContent@${String(op.at.offset)}${op.revision === undefined ? "" : "/tracked"}`;
        case DOCUMENT_OP_TYPES.DELETE_RANGE:
          return `deleteRange@${String(op.from.offset)}-${String(op.to.offset)}${op.revision === undefined ? "" : "/tracked"}`;
        default:
          return op.type;
      }
    })
    .join(", ");

// The tail batch a refused batch could not be rebased over, else the whole tail.
const refusingOps = (
  batch: DocumentBatch,
  tail: readonly SequencedBatch[],
): readonly DocumentOp[] => {
  for (const [index, remote] of tail.entries())
    if (transformBatch(batch, tail.slice(0, index + 1)).isErr())
      return remote.ops;
  return tail.flatMap(({ ops }) => ops);
};

const isTrackedDeletion = (op: DocumentOp): op is DeleteRangeOp =>
  op.type === DOCUMENT_OP_TYPES.DELETE_RANGE && op.revision !== undefined;

// Whether the batch and its tail hold tracked deletions of one paragraph.
const meetsTrackedDeletion = (
  batch: DocumentBatch,
  tail: readonly SequencedBatch[],
): boolean =>
  batch.ops.some(
    (op) =>
      isTrackedDeletion(op) &&
      tail.some(({ ops }) =>
        ops.some(
          (other) =>
            isTrackedDeletion(other) && sameParagraph(op.from, other.from),
        ),
      ),
  );

/**
 * Apply `batch` where it was authored, then each tail batch rebased over it:
 * the result must be the authority's. Undefined when a step of that order is
 * refused, which leaves nothing to compare.
 */
const commutes = (
  history: readonly Document[],
  batch: DocumentBatch,
  tail: readonly SequencedBatch[],
  result: Document,
): boolean | undefined => {
  const base = history.at(batch.baseRev);
  if (base === undefined) return undefined;
  const first = applyBatch(base, batch.ops);
  if (first.isErr()) return undefined;
  let document = first.value.document;
  let mine: SequencedBatch = {
    ...batch,
    revision: batch.baseRev + 1,
    effects: first.value.effects,
  };
  for (const remote of tail) {
    const moved = transformBatch(remote, [mine], { order: "before" });
    if (moved.isErr()) return undefined;
    const applied = applyBatch(document, moved.value.ops);
    if (applied.isErr()) return undefined;
    document = applied.value.document;
    const next = transformBatch(mine, [remote]);
    const at = history.at(remote.revision);
    if (next.isErr() || at === undefined) return undefined;
    const placed = applyBatch(at, next.value.ops);
    if (placed.isErr()) return undefined;
    mine = {
      ...next.value,
      revision: remote.revision + 1,
      effects: placed.value.effects,
    };
  }
  return equalForStaleness(document, result);
};

const simulate = (peers: number, steps: readonly Step[]): Outcome => {
  const base = canonicalBase();
  const sequencer = createSequencer(base);
  const history: Document[] = [base];
  const team: Peer[] = Array.from({ length: peers }, (_, index) => {
    const actor = `actor-${index}`;
    const granted = sequencer.grantIdLease(actor);
    if (granted.isErr()) throw granted.error;
    return {
      author: {
        actor,
        tag: String(index),
        lease: granted.value,
        allocate: createEditorIntentIdAllocator(),
      },
      client: createClient(base),
      undo: undefined,
    };
  });
  const outcome: Outcome = {
    enqueued: [],
    sequenced: [],
    rejected: [],
    dropped: [],
    documents: [],
    authority: base,
    rebased: 0,
    proven: 0,
    scopedRebased: 0,
    commentRebased: 0,
    diverged: [],
    trackedPairs: 0,
    pairs: noPairs(),
    typingRefusals: [],
    identifiedTypingRefusals: [],
    trackedOrder: noTrackedOrder(),
    trackedRefusals: [],
    replaced: noTally(),
    rebasedChanges: noTally(),
    resubmitted: 0,
  };
  const delivery: Delivery[] = [];
  // The batch each peer last submitted, which a reconnection submits again.
  const inFlight = new Map<number, DocumentBatch>();
  let serial = 10;

  const enqueue = (peer: Peer, ops: readonly DocumentOp[]): void => {
    const before = peer.client.document;
    const batch: DocumentBatch = {
      schema: DOCUMENT_OP_SCHEMA_VERSION,
      opId: `edit-${serial++}`,
      actor: peer.author.actor,
      baseRev: peer.client.headRev,
      ops,
    };
    // Every compiled edit and every undo crosses the wire.
    const sent = validateDocumentBatch(wire(batch));
    if (sent.isErr())
      throw new Error(
        `${ops.map(({ type }) => type).join(", ")} is not sendable: ${sent.error.message}`,
      );
    const queued = peer.client.enqueue(wire(batch));
    // A stale undo is dropped here, never forced.
    if (queued.isErr()) {
      expect(queued.error.reason).toBe(
        BATCH_REJECTION_REASONS.INVALID_OPERATION,
      );
      return;
    }
    outcome.enqueued.push(batch.opId);
    const applied = applyBatch(before, queued.value.ops);
    peer.undo = applied.isOk() ? applied.value.inverse : undefined;
    const minted = introducedIds(
      leasedIdCensus(before),
      leasedIdCensus(peer.client.document),
    );
    const lease = peer.author.lease;
    if (lease !== undefined)
      peer.author = { ...peer.author, lease: trimIdLease(lease, minted) };
  };

  const submit = (index: number): void => {
    const peer = team.at(index);
    const batch = peer?.client.nextSubmission();
    if (batch === undefined) return;
    const tail = sequencer.broadcasts.slice(batch.baseRev);
    const result = sequencer.submit(wire(batch));
    inFlight.set(index, batch);
    delivery.push({ peer: index, type: "result", result });
    if (result.type === "reject") {
      outcome.rejected.push(result.reason.reason);
      const { opType, overType } = result.reason;
      // Two typings refuse each other only through their insertions and
      // seam joins; another operation of the batch may refuse instead.
      const typingPair = [result.reason.opType, result.reason.overType].every(
        (type) =>
          type === DOCUMENT_OP_TYPES.INSERT_CONTENT ||
          type === DOCUMENT_OP_TYPES.JOIN_INLINE,
      );
      const refusing = refusingOps(batch, tail);
      // No declared tracked pair is refused: the authority's order decides.
      // An inverse keeps its preconditions, which refuse by their own rule,
      // an insertion against a deletion's range is the typing and deletion
      // class, and a hard inline split or join (rejoining or cutting a
      // record exactly) the record cutting class, each with rules of its own.
      for (const op of restoredOps(opsAtRefusal(batch, tail)))
        for (const other of restoredOps(refusing)) {
          const kind = trackedTransform.trackedOrderPair(op, other);
          const named =
            (op.type === opType && other.type === overType) ||
            (op.type === overType && other.type === opType);
          if (
            kind !== undefined &&
            named &&
            !holdsPreconditions(op) &&
            !holdsPreconditions(other) &&
            !insertionAgainstDeletion(op, other) &&
            !cutsRecordsExactly(op) &&
            !cutsRecordsExactly(other)
          )
            outcome.trackedRefusals.push(
              `${kind}: ${String(opType)} over ${String(overType)}: ${result.reason.message} (${describeOps(batch.ops)} over ${describeOps(refusing)})`,
            );
        }
      // An edit that cuts an identified record (a revision or a content
      // control) at the shared point, or joins its pieces again, would give
      // the record's second half the id of whichever applied first.
      const identified =
        result.reason.message.includes("identified record") ||
        [...batch.ops, ...refusing].some(
          (op) =>
            op.type === DOCUMENT_OP_TYPES.JOIN_INLINE && op.ifAlike !== true,
        );
      for (const kind of pairKinds(batch.ops, refusing)) {
        if (kind === "typingTyping" && !typingPair) continue;
        outcome.pairs[kind].refused += 1;
        if (kind !== "typingTyping") continue;
        const described = `${String(result.reason.opType)} over ${String(result.reason.overType)}: ${result.reason.message} (${describeOps(batch.ops)} over ${describeOps(refusing)})`;
        if (identified) outcome.identifiedTypingRefusals.push(described);
        else outcome.typingRefusals.push(described);
      }
      return;
    }
    history.push(sequencer.document);
    if (tail.length > 0) {
      outcome.rebased += 1;
      for (const kind of pairKinds(
        batch.ops,
        tail.flatMap(({ ops }) => ops),
      ))
        outcome.pairs[kind].accepted += 1;
      if (
        [batch, ...tail].some(({ ops }) => ops.some((op) => isBlockScoped(op)))
      )
        outcome.scopedRebased += 1;
      if ([batch, ...tail].some(({ ops }) => ops.some((op) => isCommentOp(op))))
        outcome.commentRebased += 1;
      if (meetsTrackedDeletion(batch, tail)) outcome.trackedPairs += 1;
      const proven = commutes(history, batch, tail, sequencer.document);
      if (proven === true) outcome.proven += 1;
      if (proven === false) {
        // Only a declared tracked pair may record other cuts, or leave
        // shared content to another deletion, in the other order; every
        // peer still reaches the authority's document by replay.
        const declared = declaredOver(batch, tail);
        if (declared.size === 0) outcome.diverged.push(batch.opId);
        for (const kind of declared) outcome.trackedOrder[kind] += 1;
      }
    }
    const broadcast = sequencer.broadcasts.at(-1);
    if (broadcast === undefined) return;
    for (let target = 0; target < peers; target += 1)
      delivery.push({
        peer: target,
        type: "broadcast",
        batch: wire(broadcast),
      });
  };

  const deliver = (seed: number): void => {
    const event = delivery.splice(seed % delivery.length, 1).at(0);
    const peer = event === undefined ? undefined : team.at(event.peer);
    if (event === undefined || peer === undefined) return;
    const before = peer.client.document;
    const pending = peer.client.pending.flatMap(({ ops }) => ops);
    let changes: readonly ClientDocumentChange[];
    if (event.type === "broadcast") {
      const received = peer.client.receiveBroadcast(event.batch);
      if (received.isErr()) throw received.error;
      changes = received.value;
    } else if (event.result.type === "ack") {
      peer.client.receiveAck(event.result);
      changes = [];
    } else changes = peer.client.receiveReject(event.result);
    // Every change takes the peer's document exactly where the client is:
    // a rebased one is never trusted short of it.
    let document = before;
    for (const change of changes) {
      const applied = applyBatch(document, change.batch.ops);
      if (applied.isErr()) throw applied.error;
      document = applied.value.document;
      if (change.type === "replaced" && change.moves !== undefined) {
        const moved = applyBatch(before, change.moves.ops);
        if (moved.isErr())
          throw new Error(
            `A replacement's moves do not apply where it does: ${moved.error.message}`,
          );
      }
    }
    if (!equalForStaleness(document, peer.client.document))
      throw new Error(
        `The ${changes.map(({ type }) => type).join(", ")} change of ${event.type === "broadcast" ? event.batch.opId : event.result.opId} leaves another document than the client's, first at ${JSON.stringify(differenceAt(document, peer.client.document, "document"))}`,
      );
    if (event.type !== "broadcast" || pending.length === 0) return;
    const kinds = [...pairKinds(pending, event.batch.ops)];
    for (const change of changes) {
      const tally =
        change.type === "replaced" ? outcome.replaced : outcome.rebasedChanges;
      if (kinds.length === 0) tally.other += 1;
      for (const kind of kinds) tally[kind] += 1;
    }
  };

  // The connection drops before the answer arrives: the answer is lost and
  // the batch is submitted again under its operation id, which the
  // sequencer answers as it did the first time.
  const reconnect = (index: number): void => {
    const peer = team.at(index);
    const batch = inFlight.get(index);
    if (peer === undefined || batch === undefined) return;
    if (!peer.client.pending.some(({ opId }) => opId === batch.opId)) return;
    for (let at = delivery.length - 1; at >= 0; at -= 1) {
      const event = delivery[at];
      if (
        event?.type === "result" &&
        event.peer === index &&
        event.result.opId === batch.opId
      )
        delivery.splice(at, 1);
    }
    const result = sequencer.submit(wire(batch));
    outcome.resubmitted += 1;
    delivery.push({ peer: index, type: "result", result });
  };

  for (const step of steps) {
    if (step.type === "deliver") {
      if (delivery.length > 0) deliver(step.seed);
      continue;
    }
    const index = step.peer % peers;
    const peer = team.at(index);
    if (peer === undefined) continue;
    if (step.type === "submit") {
      submit(index);
      continue;
    }
    if (step.type === "reconnect") {
      reconnect(index);
      continue;
    }
    if (step.type === "undo") {
      if (peer.undo !== undefined && peer.undo.length > 0)
        enqueue(peer, peer.undo);
      peer.undo = undefined;
      continue;
    }
    const ops = compileCanonicalEdit(
      peer.client.document,
      peer.author,
      step.seed,
      serial,
    );
    if (ops !== undefined && ops.length > 0) enqueue(peer, ops);
  }
  let rounds = 0;
  while (
    delivery.length > 0 ||
    team.some(({ client }) => client.pending.length > 0)
  ) {
    expect(rounds++).toBeLessThan(10_000);
    for (let index = 0; index < peers; index += 1) submit(index);
    if (delivery.length > 0) deliver(rounds * 7919);
  }
  return {
    ...outcome,
    sequenced: sequencer.broadcasts.map(({ opId }) => opId),
    dropped: team.flatMap(({ client }) =>
      client.notices.map(({ opId }) => opId),
    ),
    documents: team.map(({ client }) => client.document),
    authority: sequencer.document,
  };
};

const stepArbitrary = (
  peers: number,
  seeds: fc.Arbitrary<EditSeed> = editSeedArbitrary,
  undo = true,
  reconnect = false,
): fc.Arbitrary<Step> => {
  const peer = fc.nat({ max: peers - 1 });
  const editStep = {
    weight: 5,
    arbitrary: fc.record({
      type: fc.constant("edit" as const),
      peer,
      seed: seeds,
    }),
  };
  const undoStep = {
    weight: 1,
    arbitrary: fc.record({ type: fc.constant("undo" as const), peer }),
  };
  const submitStep = {
    weight: 2,
    arbitrary: fc.record({ type: fc.constant("submit" as const), peer }),
  };
  const deliverStep = {
    weight: 3,
    arbitrary: fc.record({
      type: fc.constant("deliver" as const),
      seed: fc.nat({ max: 100_000 }),
    }),
  };
  const reconnectStep = {
    weight: 1,
    arbitrary: fc.record({ type: fc.constant("reconnect" as const), peer }),
  };
  return fc.oneof(
    editStep,
    submitStep,
    deliverStep,
    ...(undo ? [undoStep] : []),
    ...(reconnect ? [reconnectStep] : []),
  );
};

type Scenario = [number, Step[]];

const scenarioArbitrary = (
  seeds?: fc.Arbitrary<EditSeed>,
  undo?: boolean,
  reconnect?: boolean,
): fc.Arbitrary<Scenario> =>
  fc.integer({ min: 3, max: 4 }).chain((peers) =>
    fc.tuple(
      fc.constant(peers),
      fc.array(stepArbitrary(peers, seeds, undo, reconnect), {
        minLength: 8,
        maxLength: 30,
      }),
    ),
  );

/** Edits of the body's first paragraph, at the same and nearby offsets. */
const focusedSeeds = (
  kinds: readonly [EditKind, ...EditKind[]],
  suggest: fc.Arbitrary<boolean>,
): fc.Arbitrary<EditSeed> =>
  fc.record({
    kind: fc.constantFrom(...kinds),
    a: fc.nat({ max: 24 }),
    b: fc.nat({ max: 100_000 }),
    suggest,
    focus: fc.constant(true as const),
  });

/** Typing and inserted atoms only, as an author typing directly does. */
const typingScenarioArbitrary = scenarioArbitrary(
  focusedSeeds(["type", "type", "tab"], fc.constant(false)),
  false,
);

/** Typing among formats, erases, comments and replies in one paragraph. */
const sameParagraphScenarioArbitrary = scenarioArbitrary(
  focusedSeeds(
    [
      "type",
      "type",
      "type",
      "tab",
      "erase",
      "replace",
      "format",
      "comment",
      "pointComment",
      "reply",
      "reply",
    ],
    fc.boolean(),
  ),
);

/**
 * Suggesting and direct typing, erases and replacements in one paragraph,
 * mostly suggesting, with lost answers submitted again. A format at a
 * tracked change's edge is pinned (`formatAcrossSuggestedErase`).
 */
const trackedScenarioArbitrary = scenarioArbitrary(
  focusedSeeds(
    ["type", "type", "erase", "erase", "replace", "tab"],
    fc.constantFrom(true, true, false),
  ),
  true,
  true,
);

const bytes = async (document: Document): Promise<Uint8Array> =>
  new Uint8Array(await serializeDocumentToDocx(document));

type Counts = {
  rebased: number;
  proven: number;
  scopedRebased: number;
  commentRebased: number;
  trackedPairs: number;
  pairs: PairCounts;
  /** Why two typings refused each other: the pair's types and message. */
  typingRefusals: string[];
  /** Typing pairs refused because one side cuts or rejoins an identified record. */
  identifiedTypingRefusals: string[];
  trackedOrder: Record<trackedTransform.TrackedOrderPair, number>;
  replaced: Tally;
  rebasedChanges: Tally;
  resubmitted: number;
};

const noCounts = (): Counts => ({
  rebased: 0,
  proven: 0,
  scopedRebased: 0,
  commentRebased: 0,
  trackedPairs: 0,
  pairs: noPairs(),
  typingRefusals: [],
  identifiedTypingRefusals: [],
  trackedOrder: noTrackedOrder(),
  replaced: noTally(),
  rebasedChanges: noTally(),
  resubmitted: 0,
});

const convergence = (
  counts: Counts = noCounts(),
  scenarios: fc.Arbitrary<Scenario> = scenarioArbitrary(),
  // Whether every edit must be accepted, rather than refused with a type.
  acceptAll = false,
) =>
  fc.asyncProperty(scenarios, async ([peers, steps]) => {
    const outcome = simulate(peers, steps);
    counts.rebased += outcome.rebased;
    counts.proven += outcome.proven;
    counts.scopedRebased += outcome.scopedRebased;
    counts.commentRebased += outcome.commentRebased;
    counts.trackedPairs += outcome.trackedPairs;
    for (const kind of PAIR_KINDS) {
      counts.pairs[kind].accepted += outcome.pairs[kind].accepted;
      counts.pairs[kind].refused += outcome.pairs[kind].refused;
    }
    counts.typingRefusals.push(...outcome.typingRefusals);
    counts.identifiedTypingRefusals.push(...outcome.identifiedTypingRefusals);
    for (const kind of TRACKED_ORDER_KINDS)
      counts.trackedOrder[kind] += outcome.trackedOrder[kind];
    for (const kind of [...PAIR_KINDS, "other"] as const) {
      counts.replaced[kind] += outcome.replaced[kind];
      counts.rebasedChanges[kind] += outcome.rebasedChanges[kind];
    }
    counts.resubmitted += outcome.resubmitted;
    // No declared tracked pair is refused.
    expect(outcome.trackedRefusals).toEqual([]);
    // Every refusal is typed, and none is a wire, base or lease failure.
    expect(outcome.rejected.filter((reason) => FORBIDDEN.has(reason))).toEqual(
      [],
    );
    if (acceptAll) {
      expect(outcome.rejected).toEqual([]);
      expect(outcome.dropped).toEqual([]);
    }
    // Concurrent batches the authority accepted commute with what they were
    // rebased over: the transform never accepts a pair it maps wrongly.
    if (outcome.diverged.length > 0)
      throw new Error(
        `Accepted batches do not commute with their tail: ${outcome.diverged.join(", ")}`,
      );
    // No acknowledged edit is lost, and no edit vanishes unreported.
    expect(new Set(outcome.sequenced).size).toBe(outcome.sequenced.length);
    expect([...outcome.sequenced, ...outcome.dropped].toSorted()).toEqual(
      outcome.enqueued.toSorted(),
    );
    const authority = await bytes(outcome.authority);
    for (const document of outcome.documents) {
      expect(equalForStaleness(document, outcome.authority)).toBe(true);
      expect(await bytes(document)).toEqual(authority);
    }
  });

const edit = (peer: number, seed: Omit<EditSeed, "suggest">): Step => ({
  type: "edit",
  peer,
  seed: { ...seed, suggest: false },
});

/** Three peers restructure one table at once: rows at different indexes. */
const concurrentRows: Scenario = [
  3,
  [
    edit(0, { kind: "insertRow", a: 1, b: 0 }),
    edit(1, { kind: "insertRow", a: 2, b: 3 }),
    edit(2, { kind: "type", a: 0, b: 0 }),
    { type: "submit", peer: 0 },
    { type: "submit", peer: 1 },
    { type: "submit", peer: 2 },
  ],
];

// The queue drains in order: every answer and broadcast reaches its peer.
const drain = (count: number): Step[] =>
  Array.from({ length: count }, () => ({ type: "deliver", seed: 0 }));

/**
 * A comment on the first paragraph's fourth character while another peer
 * erases its first: the comment's range moves back by one.
 */
const commentOverErase: Scenario = [
  3,
  [
    edit(0, { kind: "erase", a: 0, b: 0 }),
    edit(1, { kind: "comment", a: 24, b: 0 }),
    { type: "submit", peer: 0 },
    { type: "submit", peer: 1 },
  ],
];

/**
 * One thread, seen by every peer, then replied to, resolved and edited at
 * once: the three commute.
 */
const concurrentThread: Scenario = [
  3,
  [
    edit(0, { kind: "comment", a: 24, b: 0 }),
    { type: "submit", peer: 0 },
    ...drain(4),
    edit(1, { kind: "reply", a: 0, b: 0 }),
    edit(2, { kind: "resolveComment", a: 0, b: 0 }),
    edit(0, { kind: "editComment", a: 0, b: 0 }),
    { type: "submit", peer: 1 },
    { type: "submit", peer: 2 },
    { type: "submit", peer: 0 },
  ],
];

/** A thread deleted while another peer resolves it: the later one refuses. */
const deleteAgainstResolve: Scenario = [
  3,
  [
    edit(0, { kind: "pointComment", a: 36, b: 0 }),
    { type: "submit", peer: 0 },
    ...drain(4),
    edit(1, { kind: "deleteComment", a: 0, b: 0 }),
    edit(2, { kind: "resolveComment", a: 0, b: 0 }),
    { type: "undo", peer: 1 },
    { type: "submit", peer: 1 },
    { type: "submit", peer: 2 },
  ],
];

/**
 * Two suggesting peers erase one character each in the fifth body
 * paragraph, the second ending where the first starts: the tracked
 * deletions are adjacent, and whichever wraps first records the cut between
 * them.
 */
const adjacentSuggestedErases: Scenario = [
  4,
  [
    {
      type: "edit",
      peer: 0,
      seed: { kind: "erase", a: 47_038, b: 0, suggest: true },
    },
    {
      type: "edit",
      peer: 1,
      seed: { kind: "erase", a: 83_446, b: 0, suggest: true },
    },
    { type: "deliver", seed: 0 },
    edit(0, { kind: "type", a: 0, b: 0 }),
    { type: "deliver", seed: 0 },
    edit(0, { kind: "type", a: 0, b: 0 }),
    edit(0, { kind: "type", a: 0, b: 0 }),
    { type: "undo", peer: 0 },
    edit(0, { kind: "type", a: 0, b: 0 }),
    edit(0, { kind: "type", a: 0, b: 0 }),
  ],
];

/** A suggesting erase of the body's first paragraph from `a`. */
const suggestedErase = (
  peer: number,
  seed: Omit<EditSeed, "kind" | "focus" | "suggest">,
): Step => ({
  type: "edit",
  peer,
  seed: { ...seed, kind: "erase", suggest: true, focus: true },
});

const submitAll: Step[] = [0, 1, 2].map((peer) => ({ type: "submit", peer }));

/**
 * Concurrent erases in the first paragraph, each phase drained before the
 * next: a tracked deletion of 2..5 with one nested in it (3..4) and one
 * adjacent to it (5..7); the same range twice (10..12) beside a separate one
 * (14..15); a tracked deletion of 6..9 with a direct one nested in it (7..8).
 */
const concurrentSuggestedErases: Scenario = [
  3,
  [
    suggestedErase(0, { a: 2, b: 2 }),
    suggestedErase(1, { a: 3, b: 0 }),
    suggestedErase(2, { a: 5, b: 1 }),
    ...submitAll,
    ...drain(24),
    suggestedErase(0, { a: 10, b: 1 }),
    suggestedErase(1, { a: 10, b: 1 }),
    suggestedErase(2, { a: 14, b: 0 }),
    ...submitAll,
    ...drain(24),
    suggestedErase(0, { a: 6, b: 2 }),
    {
      type: "edit",
      peer: 1,
      seed: { kind: "erase", a: 7, b: 0, suggest: false, focus: true },
    },
    ...submitAll,
  ],
];

/** An edit of the body's first paragraph at `a`. */
const focused = (
  peer: number,
  seed: Omit<EditSeed, "focus" | "suggest">,
): Step => ({
  type: "edit",
  peer,
  seed: { ...seed, suggest: false, focus: true },
});

/**
 * Three peers in the first paragraph ("Paragraph 0 text"), each phase
 * concurrent and drained before the next:
 * 1. typing at one offset (4, twice) and the next one (5);
 * 2. typing at 8 inside a concurrent format of 6..9, beside a concurrent
 *    deletion of 9..11 that the format ends against;
 * 3. a comment over the text typed in phase 1 (4..6), with typing inside it
 *    (5) and after its reference (8);
 * 4. two replies to that comment.
 */
const sameParagraphTyping: Scenario = [
  3,
  [
    focused(0, { kind: "type", a: 4, b: 0 }),
    focused(1, { kind: "type", a: 4, b: 1 }),
    focused(2, { kind: "type", a: 5, b: 0 }),
    ...submitAll,
    ...drain(24),
    focused(2, { kind: "type", a: 8, b: 1 }),
    focused(0, { kind: "format", a: 6, b: 2 }),
    focused(1, { kind: "erase", a: 9, b: 1 }),
    { type: "submit", peer: 2 },
    { type: "submit", peer: 0 },
    { type: "submit", peer: 1 },
    ...drain(24),
    focused(0, { kind: "comment", a: 4, b: 1 }),
    focused(1, { kind: "type", a: 5, b: 0 }),
    focused(2, { kind: "type", a: 8, b: 0 }),
    ...submitAll,
    ...drain(24),
    focused(1, { kind: "reply", a: 0, b: 0 }),
    focused(2, { kind: "reply", a: 0, b: 0 }),
    { type: "submit", peer: 1 },
    { type: "submit", peer: 2 },
  ],
];
const SAME_PARAGRAPH_EDITS = sameParagraphTyping[1].filter(
  ({ type }) => type === "edit",
).length;

/**
 * A suggesting erase of 2..4 in the first paragraph while another peer
 * formats 3..6: the format cuts the run inside the tracked range, which
 * decides whether one record ran across all of it.
 */
const formatAcrossSuggestedErase: Scenario = [
  3,
  [
    suggestedErase(0, { a: 2, b: 1 }),
    {
      type: "edit",
      peer: 1,
      seed: { kind: "format", a: 3, b: 2, suggest: false, focus: true },
    },
    { type: "submit", peer: 0 },
    { type: "submit", peer: 1 },
  ],
];

/**
 * Two peers accept the same suggestion at once and the one sequenced first
 * undoes its decision: the undo puts back the change the other peer decided.
 */
const undoAgainstRepeatedDecision: Scenario = [
  3,
  [
    edit(1, { kind: "decide", a: 0, b: 0 }),
    edit(0, { kind: "decide", a: 0, b: 0 }),
    { type: "undo", peer: 0 },
    { type: "deliver", seed: 0 },
    edit(0, { kind: "type", a: 0, b: 0 }),
    { type: "deliver", seed: 0 },
    edit(0, { kind: "type", a: 0, b: 0 }),
    edit(0, { kind: "type", a: 0, b: 0 }),
    { type: "deliver", seed: 0 },
    { type: "submit", peer: 0 },
    { type: "submit", peer: 0 },
  ],
];

/** A suggesting edit of the body's first paragraph at `a`. */
const suggested = (
  peer: number,
  seed: Omit<EditSeed, "focus" | "suggest">,
): Step => ({
  type: "edit",
  peer,
  seed: { ...seed, suggest: true, focus: true },
});

/**
 * Tracked changes meeting in the first paragraph, each phase concurrent:
 * 1. peer 0 suggests erasing 4..5 and typing at 10, both pending, while
 *    peer 1 suggests erasing 5..6, sequenced first; both lose their first
 *    answer and submit again; peer 2 has nothing pending;
 * 2. peer 0 suggests erasing 12..15, peer 1 erases 14..17 directly (one
 *    edge of it), peer 2 suggests erasing 13..14 (inside it).
 */
const trackedPending: Scenario = [
  3,
  [
    suggestedErase(0, { a: 4, b: 0 }),
    suggested(0, { kind: "type", a: 10, b: 0 }),
    suggestedErase(1, { a: 5, b: 0 }),
    { type: "submit", peer: 1 },
    { type: "reconnect", peer: 1 },
    { type: "submit", peer: 0 },
    { type: "deliver", seed: 0 },
    { type: "reconnect", peer: 0 },
    ...drain(24),
    { type: "submit", peer: 0 },
    ...drain(24),
    suggestedErase(0, { a: 12, b: 2 }),
    focused(1, { kind: "erase", a: 14, b: 2 }),
    suggestedErase(2, { a: 13, b: 0 }),
    ...submitAll,
  ],
];

/**
 * A direct format of 10..12 in the first paragraph's one run while another
 * peer erases 9..12 directly: the format cut the run, and the erase rebased
 * over it joins the run's pieces again, so both orders leave one run.
 */
const formatInsideErase: Scenario = [
  3,
  [
    {
      type: "edit",
      peer: 0,
      seed: { kind: "format", a: 10, b: 1, suggest: false, focus: true },
    },
    {
      type: "edit",
      peer: 1,
      seed: { kind: "erase", a: 9, b: 2, suggest: false, focus: true },
    },
    { type: "submit", peer: 0 },
    { type: "submit", peer: 1 },
  ],
];

/** The pinned scenarios, which the fixed coverage sample starts with. */
const PINNED: readonly Scenario[] = [
  concurrentRows,
  commentOverErase,
  concurrentThread,
  deleteAgainstResolve,
  adjacentSuggestedErases,
  concurrentSuggestedErases,
  formatAcrossSuggestedErase,
  undoAgainstRepeatedDecision,
  trackedPending,
  formatInsideErase,
];

describe("canonical sequencing", () => {
  test("a format inside a concurrent direct erase leaves one run in either order", () => {
    const outcome = simulate(...formatInsideErase);
    expect(outcome.rejected).toEqual([]);
    expect(outcome.dropped).toEqual([]);
    expect(outcome.diverged).toEqual([]);
    expect(outcome.rebased).toBeGreaterThan(0);
    expect(outcome.proven).toBe(outcome.rebased);
    for (const document of outcome.documents)
      expect(equalForStaleness(document, outcome.authority)).toBe(true);
  });

  test(
    "a format effect without its cut run is caught",
    async () => {
      // The test-only break: sequenced formats record no cut run, so the
      // erase rebased over one leaves the run cut.
      const original = { ...applyBatchModule };
      mock.module("../applyBatch", () => ({
        ...original,
        applyBatch: (...args: Parameters<typeof original.applyBatch>) =>
          original.applyBatch(...args).map((applied) =>
            Object.assign({}, applied, {
              effects: applied.effects.map((effect) =>
                effect.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS
                  ? { type: DOCUMENT_OP_TYPES.SET_RUN_PROPS }
                  : effect,
              ),
            }),
          ),
      }));
      try {
        const result = await fc.check(
          convergence(),
          // The break is a counterexample on purpose: report nothing.
          propertyConfig({
            numRuns: 1,
            examples: [[formatInsideErase]],
            reporter: () => undefined,
          }),
        );
        expect(result.failed).toBe(true);
        expect(String(result.errorInstance)).toMatch(/do not commute/u);
      } finally {
        mock.module("../applyBatch", () => original);
      }
    },
    propertyTestTimeout(60_000),
  );

  test(
    "concurrent canonical edits converge on the authority's bytes and commute where accepted",
    async () => {
      await assertProperty(convergence(), {
        numRuns: NUM_RUNS,
        id: "canonical sequencing convergence",
        examples: PINNED.map((scenario) => [scenario]),
      });
      // On a fixed sample, which starts with the pinned scenarios, the
      // schedules race: accepted batches are rebased over concurrent ones,
      // block-scoped, comment and tracked deletion ones among them, and the
      // commuted order is proven for them.
      const counts = noCounts();
      await checkCoverageSample(convergence(counts), NUM_RUNS + PINNED.length, {
        map: (value, index) => {
          const pinned = PINNED[index];
          return pinned === undefined ? value : [pinned];
        },
      });
      expect(counts.rebased).toBeGreaterThan(0);
      expect(counts.scopedRebased).toBeGreaterThan(0);
      expect(counts.commentRebased).toBeGreaterThan(0);
      expect(counts.proven).toBeGreaterThan(0);
      expect(counts.trackedPairs).toBeGreaterThan(0);
    },
    propertyTestTimeout(180_000),
  );

  test("concurrent tracked changes that meet are accepted in the authority's order", () => {
    const accepted = (scenario: Scenario) => {
      const outcome = simulate(...scenario);
      // Nothing is refused or dropped, every batch the authority accepted
      // commutes or forms a declared tracked pair, and every peer ends on
      // the authority's document.
      expect(outcome.rejected).toEqual([]);
      expect(outcome.dropped).toEqual([]);
      expect(outcome.diverged).toEqual([]);
      for (const document of outcome.documents)
        expect(equalForStaleness(document, outcome.authority)).toBe(true);
      return outcome;
    };
    // Adjacent tracked deletions: whichever wraps first records the cut
    // between them.
    expect(
      accepted(adjacentSuggestedErases).trackedOrder.cutAtTrackedChange,
    ).toBeGreaterThan(0);
    // Nested, adjacent and equal ranges: the deletion sequenced first holds
    // the shared content.
    const phases = accepted(concurrentSuggestedErases);
    expect(phases.trackedOrder.overlappingDeletions).toBeGreaterThan(0);
    expect(phases.trackedPairs).toBeGreaterThan(0);
    // A format cutting runs inside a tracked range.
    expect(
      accepted(formatAcrossSuggestedErase).trackedOrder.cutAtTrackedChange,
    ).toBeGreaterThan(0);
    // A peer with two pending batches gets the authority's document as a
    // replacement; answers lost to a reconnection are answered again.
    const pending = accepted(trackedPending);
    expect(pending.replaced.suggestedDeletions).toBeGreaterThan(0);
    expect(pending.resubmitted).toBe(2);
    expect(pending.trackedOrder.cutAtTrackedChange).toBeGreaterThan(0);
    expect(pending.trackedOrder.deletionAtTrackedEdge).toBeGreaterThan(0);
  });

  test(
    "concurrent tracked changes in one paragraph are never refused for the cuts they record",
    async () => {
      // The bar: no declared tracked pair is refused (checked for every
      // scenario); other refusals stay typed.
      await assertProperty(convergence(noCounts(), trackedScenarioArbitrary), {
        numRuns: NUM_RUNS,
        id: "canonical sequencing tracked pairs",
        examples: [[trackedPending]],
      });
      const counts = noCounts();
      await checkCoverageSample(
        convergence(counts, trackedScenarioArbitrary),
        NUM_RUNS + 2,
        {
          map: (value, index) => {
            const pinned = [trackedPending, concurrentSuggestedErases][index];
            return pinned === undefined ? value : [pinned];
          },
        },
      );
      expect(
        TRACKED_ORDER_KINDS.filter((kind) => counts.trackedOrder[kind] === 0),
      ).toEqual([]);
      expect(counts.pairs.suggestedDeletions.accepted).toBeGreaterThan(0);
      expect(counts.pairs.suggestedEdits.accepted).toBeGreaterThan(0);
      expect(counts.replaced.suggestedDeletions).toBeGreaterThan(0);
      expect(counts.resubmitted).toBeGreaterThan(0);
    },
    propertyTestTimeout(180_000),
  );

  test("an undo that puts back a change a concurrent peer decided is refused", () => {
    const outcome = simulate(...undoAgainstRepeatedDecision);
    expect(outcome.diverged).toEqual([]);
    expect(outcome.rejected).toContain(BATCH_REJECTION_REASONS.PART_CONFLICT);
    for (const document of outcome.documents)
      expect(equalForStaleness(document, outcome.authority)).toBe(true);
  });

  test(
    "an undo let through over a concurrent decision of its change is caught",
    async () => {
      // The test-only break: no operation holds a revision, so an undo that
      // puts a change back passes a concurrent decision of it.
      const original = { ...heldRevisions };
      mock.module("../heldRevisions", () => ({
        ...original,
        heldRevisionIds: () => [],
      }));
      try {
        const result = await fc.check(
          convergence(),
          // The break is a counterexample on purpose: report nothing.
          propertyConfig({
            numRuns: NUM_RUNS,
            examples: [[undoAgainstRepeatedDecision]],
            reporter: () => undefined,
          }),
        );
        expect(result.failed).toBe(true);
        // Caught by the commutation check, not by anything else.
        expect(String(result.errorInstance)).toMatch(/do not commute/u);
      } finally {
        mock.module("../heldRevisions", () => original);
      }
      // Without the break the same scenario passes.
      const restored = await fc.check(
        convergence(),
        propertyConfig({
          numRuns: 1,
          examples: [[undoAgainstRepeatedDecision]],
        }),
      );
      expect(restored.failed).toBe(false);
    },
    propertyTestTimeout(180_000),
  );

  test(
    "concurrent typing in one paragraph is never refused",
    async () => {
      // Every typed edit is accepted: the bar for typing pairs is no refusal.
      await assertProperty(
        convergence(noCounts(), typingScenarioArbitrary, true),
        { numRuns: NUM_RUNS, id: "canonical sequencing concurrent typing" },
      );
      const counts = noCounts();
      await checkCoverageSample(
        convergence(counts, typingScenarioArbitrary, true),
        NUM_RUNS,
      );
      expect(counts.pairs.typingTyping.accepted).toBeGreaterThan(0);
      expect(counts.typingRefusals).toEqual([]);
      expect(counts.proven).toBeGreaterThan(0);
    },
    propertyTestTimeout(180_000),
  );

  test(
    "typing among formats, erases, comments and replies in one paragraph converges and commutes where accepted",
    async () => {
      await assertProperty(
        convergence(noCounts(), sameParagraphScenarioArbitrary),
        {
          numRuns: NUM_RUNS,
          id: "canonical sequencing same-paragraph convergence",
          examples: [[sameParagraphTyping]],
        },
      );
      // On a fixed sample, which starts with the pinned scenario, typing is
      // rebased over every kind of concurrent edit of its paragraph, and two
      // replies to one parent over each other; typing pairs are never
      // refused, unless one side cuts or rejoins an identified record at the
      // shared point.
      const counts = noCounts();
      await checkCoverageSample(
        convergence(counts, sameParagraphScenarioArbitrary),
        NUM_RUNS + 1,
        {
          map: (value, index) => (index === 0 ? [sameParagraphTyping] : value),
        },
      );
      expect(
        TYPING_PAIR_KINDS.filter((kind) => counts.pairs[kind].accepted === 0),
      ).toEqual([]);
      expect(counts.typingRefusals).toEqual([]);
    },
    propertyTestTimeout(180_000),
  );

  test("concurrent typing, formatting, deletion, comments and replies in one paragraph are never refused", () => {
    const outcome = simulate(...sameParagraphTyping);
    // Every edit compiled, none was refused or dropped, and each pair the
    // scenario sets up was rebased, accepted and proven to commute.
    expect(outcome.enqueued.length).toBe(SAME_PARAGRAPH_EDITS);
    expect(outcome.rejected).toEqual([]);
    expect(outcome.dropped).toEqual([]);
    expect(outcome.diverged).toEqual([]);
    expect(
      TYPING_PAIR_KINDS.filter((kind) => outcome.pairs[kind].accepted === 0),
    ).toEqual([]);
    for (const document of outcome.documents)
      expect(equalForStaleness(document, outcome.authority)).toBe(true);
  });

  test(
    "an inverted tie between concurrent insertions at one point is caught",
    async () => {
      // The test-only break: an inline insertion pair always takes the
      // later side's tie-break, so at one gap both orders put the other
      // insertion first.
      const original = { ...inlineTransform };
      mock.module("../inlineTransform", () => ({
        ...original,
        transformInlinePair: (
          pair: Parameters<typeof original.transformInlinePair>[0],
        ) => original.transformInlinePair({ ...pair, order: "after" }),
      }));
      try {
        const result = await fc.check(
          convergence(),
          // The break is a counterexample on purpose: report nothing.
          propertyConfig({
            numRuns: NUM_RUNS,
            examples: [[sameParagraphTyping]],
            reporter: () => undefined,
          }),
        );
        expect(result.failed).toBe(true);
        // Caught by the commutation check, not by anything else.
        expect(String(result.errorInstance)).toMatch(/do not commute/u);
      } finally {
        mock.module("../inlineTransform", () => original);
      }
      // Without the break the same scenario passes.
      const restored = await fc.check(
        convergence(),
        propertyConfig({ numRuns: 1, examples: [[sameParagraphTyping]] }),
      );
      expect(restored.failed).toBe(false);
    },
    propertyTestTimeout(180_000),
  );

  test(
    "a non-commuting pair outside the declared tracked kinds is caught",
    async () => {
      // The test-only break: no pair is declared order-dependent, so the two
      // adjacent suggesting erases, accepted in the authority's order, no
      // longer have a declared reason to differ in the other order.
      const original = { ...trackedTransform };
      mock.module("../trackedTransform", () => ({
        ...original,
        trackedOrderPair: () => undefined,
      }));
      try {
        const result = await fc.check(
          convergence(),
          // The break is a counterexample on purpose: report nothing.
          propertyConfig({
            numRuns: NUM_RUNS,
            examples: [[adjacentSuggestedErases]],
            reporter: () => undefined,
          }),
        );
        expect(result.failed).toBe(true);
        // Caught by the commutation check, not by anything else.
        expect(String(result.errorInstance)).toMatch(/do not commute/u);
      } finally {
        mock.module("../trackedTransform", () => original);
      }
      // Without the break the same scenario passes.
      const restored = await fc.check(
        convergence(),
        propertyConfig({ numRuns: 1, examples: [[adjacentSuggestedErases]] }),
      );
      expect(restored.failed).toBe(false);
    },
    propertyTestTimeout(180_000),
  );

  test(
    "a transform that skips the table structure check is caught",
    async () => {
      // The test-only break: no operation counts as a table structure edit,
      // so the footprint check lets two edits of one table's rows through.
      const original = { ...envelope };
      mock.module("../envelope", () => ({
        ...original,
        TABLE_EXCLUSIVE_OP_TYPES: {},
      }));
      try {
        const result = await fc.check(
          convergence(),
          // The break is a counterexample on purpose: report nothing.
          propertyConfig({
            numRuns: NUM_RUNS,
            examples: [[concurrentRows]],
            reporter: () => undefined,
          }),
        );
        expect(result.failed).toBe(true);
        // Caught by the commutation check, not by anything else.
        expect(String(result.errorInstance)).toMatch(/do not commute/u);
      } finally {
        mock.module("../envelope", () => original);
      }
      // Without the break the same scenario passes.
      const restored = await fc.check(
        convergence(),
        propertyConfig({ numRuns: 1, examples: [[concurrentRows]] }),
      );
      expect(restored.failed).toBe(false);
    },
    propertyTestTimeout(180_000),
  );

  test(
    "a comment anchor left in place over a concurrent deletion is caught",
    async () => {
      // The test-only break: a comment's anchor no longer moves over a
      // concurrent text or structure operation.
      const original = { ...commentTransform };
      mock.module("../commentTransform", () => ({
        ...original,
        moveCommentAnchor: ({ op }: { op: DocumentOp }) => Result.ok([op]),
      }));
      try {
        const result = await fc.check(
          convergence(),
          // The break is a counterexample on purpose: report nothing.
          propertyConfig({
            numRuns: NUM_RUNS,
            examples: [[commentOverErase]],
            reporter: () => undefined,
          }),
        );
        expect(result.failed).toBe(true);
        // Caught by the commutation check, not by anything else.
        expect(String(result.errorInstance)).toMatch(/do not commute/u);
      } finally {
        mock.module("../commentTransform", () => original);
      }
      // Without the break the same scenario passes.
      const restored = await fc.check(
        convergence(),
        propertyConfig({ numRuns: 1, examples: [[commentOverErase]] }),
      );
      expect(restored.failed).toBe(false);
    },
    propertyTestTimeout(180_000),
  );

  test(
    "a comment id outside the submitter's lease is refused",
    () => {
      assertProperty(
        fc.property(
          fc.nat({ max: 1023 }),
          fc.constantFrom("comment" as const, "pointComment" as const),
          (offset, kind) => {
            const base = canonicalBase();
            const sequencer = createSequencer(base);
            const own = sequencer.grantIdLease("actor-0");
            const other = sequencer.grantIdLease("actor-1");
            if (own.isErr()) throw own.error;
            if (other.isErr()) throw other.error;
            const block = other.value.blocks.find(
              ({ kind: leased }) => leased === ID_LEASE_KINDS.COMMENT,
            );
            if (block === undefined) throw new Error("No comment lease block.");
            const ops = compileCanonicalEdit(
              base,
              {
                actor: "actor-0",
                tag: "0",
                lease: own.value,
                allocate: createEditorIntentIdAllocator(),
              },
              { kind, a: 24, b: 0, suggest: false },
              1,
            );
            const created = ops?.at(0);
            if (created?.type !== DOCUMENT_OP_TYPES.CREATE_COMMENT)
              throw new Error("The comment edit did not compile.");
            // The same comment under an id another actor leases.
            const forged = {
              ...created,
              comment: {
                ...created.comment,
                id: block.start + (offset % (block.end - block.start)),
              },
            };
            const result = sequencer.submit(
              wire({
                schema: DOCUMENT_OP_SCHEMA_VERSION,
                opId: "forged",
                actor: "actor-0",
                baseRev: 0,
                ops: [forged],
              }),
            );
            expect(result.type).toBe("reject");
            if (result.type !== "reject") return;
            expect(result.reason.reason).toBe(
              BATCH_REJECTION_REASONS.OUTSIDE_ID_LEASE,
            );
            expect(result.reason.idKind).toBe(ID_LEASE_KINDS.COMMENT);
            // The honest comment is accepted.
            const honest = sequencer.submit(
              wire({
                schema: DOCUMENT_OP_SCHEMA_VERSION,
                opId: "honest",
                actor: "actor-0",
                baseRev: 0,
                ops: [created],
              }),
            );
            expect(honest.type).toBe("ack");
          },
        ),
        { numRuns: 20, id: "comment ids outside the submitter's lease" },
      );
    },
    propertyTestTimeout(60_000),
  );
});
