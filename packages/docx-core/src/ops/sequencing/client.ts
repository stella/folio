import { panic, Result } from "better-result";

import type { Document } from "../../model/document";
import { storyParagraphs } from "../blocks";
import { validateOpsDocument } from "../contract";
import { equalForStaleness, structurallyEqual } from "../equality";
import { paragraphLength } from "../offsets";
import { documentStories, findStoryBody } from "../stories";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
} from "../types";
import { captureDocumentOp, restoreDocumentOp } from "../wire";
import { applyBatch } from "./applyBatch";
import {
  BATCH_REJECTION_REASONS,
  BatchRejection,
  type DocumentBatch,
  type SequencedBatch,
  validateDocumentBatch,
  validateSequencedBatch,
} from "./envelope";
import type { BatchAck, BatchReject } from "./sequencer";
import { transformBatch } from "./transform";

const PENDING_STATUSES = {
  QUEUED: "queued",
  SUBMITTED: "submitted",
  ACKNOWLEDGED: "acknowledged",
} as const;

type PendingEntry = {
  batch: DocumentBatch;
  inverse: readonly DocumentOp[];
  effects: NonNullable<SequencedBatch["effects"]>;
  status: (typeof PENDING_STATUSES)[keyof typeof PENDING_STATUSES];
};

export type ClientNotice = {
  opId: string;
  reason: BatchRejection;
  ops: readonly DocumentOp[];
};

/**
 * How the optimistic `document` moved when the authority answered: applied in
 * order to the document before the answer, `batch.ops` give the one after it.
 * `rebased`: a foreign revision rebased over every pending local batch, or a
 * refused local batch's inverse rebased over the ones after it, so positions
 * and history can be transformed over it. `replaced`: a pending batch could not
 * be rebased, or the rebased change would not reach the replayed document, so
 * the change undoes every pending batch, applies what the authority sequenced
 * and replays the survivors; only the document is exact. When every pending
 * batch survived and the rebased change reaches a document with the same
 * paragraphs at the same lengths (the two differ only in how records are cut
 * or which tracked change holds some content), `moves` is that rebased
 * change: positions move over it exactly as over a `rebased` one.
 */
export type ClientDocumentChange =
  | { type: "rebased"; batch: SequencedBatch }
  | { type: "replaced"; batch: SequencedBatch; moves?: SequencedBatch };

// Whether two documents hold the same paragraphs, story by story, at the same
// lengths: a position by paragraph id and offset means the same place in both.
const sameOffsets = (left: Document, right: Document): boolean => {
  const shape = (document: Document) =>
    documentStories(document).map((story) => {
      const body = findStoryBody(document, story);
      return body === undefined
        ? []
        : storyParagraphs(body).map(({ paragraph }) => [
            paragraph.paraId,
            paragraphLength(paragraph),
          ]);
    });
  return structurallyEqual(shape(left), shape(right));
};

/** Optimistic state with ordered authoritative replay and one causal submission in flight. */
export const createClient = (
  document: Document,
): {
  enqueue: (batch: DocumentBatch) => Result<DocumentBatch, BatchRejection>;
  nextSubmission: () => DocumentBatch | undefined;
  receiveAck: (ack: BatchAck) => void;
  receiveReject: (rejection: BatchReject) => readonly ClientDocumentChange[];
  receiveBroadcast: (
    batch: SequencedBatch,
  ) => Result<readonly ClientDocumentChange[], BatchRejection>;
  readonly document: Document;
  readonly headRev: number;
  readonly pending: readonly DocumentBatch[];
  readonly notices: readonly ClientNotice[];
} => {
  const valid = validateOpsDocument(document);
  if (valid.isErr()) {
    throw valid.error;
  }
  let confirmed = document;
  let current = document;
  let headRev = 0;
  let entries: PendingEntry[] = [];
  let inFlight: string | undefined;
  const received = new Map<number, SequencedBatch>();
  const history = new Map<number, SequencedBatch>();
  const appliedIds = new Set<string>();
  const localIds = new Set<string>();
  const notices: ClientNotice[] = [];

  const drop = (entry: PendingEntry, reason: BatchRejection): void => {
    notices.push({ opId: entry.batch.opId, reason, ops: entry.batch.ops });
  };

  // Every pending batch undone, the sequenced ops applied, the survivors replayed.
  const replacement = (
    previous: Document,
    previousEntries: readonly PendingEntry[],
    sequenced: readonly DocumentOp[],
    source: Pick<SequencedBatch, "opId" | "actor" | "revision">,
    moves?: SequencedBatch,
  ): ClientDocumentChange => {
    const ops = [
      ...[...previousEntries].reverse().flatMap(({ inverse }) => inverse),
      ...sequenced,
      ...entries.flatMap(({ batch }) => batch.ops),
    ];
    const applied = applyBatch(previous, ops);
    if (applied.isErr())
      return panic(
        "Pending inverses, sequenced ops and survivors must replay in order.",
      );
    return {
      type: "replaced",
      batch: {
        schema: DOCUMENT_OP_SCHEMA_VERSION,
        ...source,
        baseRev: source.revision - 1,
        ops,
        effects: applied.value.effects,
      },
      ...(moves === undefined ? {} : { moves }),
    };
  };

  // A rebased change is trusted only when it lands where the replay did: a
  // transform that does not commute for some pair then costs the history,
  // never a document that differs from the authority's. Where it lands
  // short of the replay, positions still follow it when the two documents
  // hold the same paragraphs at the same lengths.
  const landing = (
    previous: Document,
    change: SequencedBatch,
  ): "exact" | "offsets" | "elsewhere" => {
    const applied = applyBatch(previous, change.ops);
    if (applied.isErr()) return "elsewhere";
    if (equalForStaleness(applied.value.document, current)) return "exact";
    return sameOffsets(applied.value.document, current)
      ? "offsets"
      : "elsewhere";
  };

  // A batch's ops applied in turn, or undefined when one does not apply.
  const documentAfter = (
    start: Document,
    batches: readonly SequencedBatch[],
  ): Document | undefined => {
    let cursor = start;
    for (const { ops } of batches) {
      const applied = applyBatch(cursor, ops);
      if (applied.isErr()) return undefined;
      cursor = applied.value.document;
    }
    return cursor;
  };

  /**
   * One pending batch rebased over `leading`, the batches leading from the
   * document it was made in (`at`) to the one it now applies to, and
   * `leading` carried past it. When `leading` does not carry past it, the
   * later batches move over its inverse, `leading` and its rebased self.
   */
  const passOver = (
    entry: PendingEntry,
    leading: readonly SequencedBatch[],
    at: Document,
  ): Result<
    { batch: DocumentBatch; leading: SequencedBatch[] },
    BatchRejection
  > => {
    const mine: SequencedBatch = {
      ...entry.batch,
      revision: headRev + 1,
      effects: entry.effects,
    };
    const moved = transformBatch(entry.batch, leading);
    if (moved.isErr()) return Result.err(moved.error);
    const carried: SequencedBatch[] = [];
    let cursor = at;
    let carriedMine: SequencedBatch = mine;
    for (const lead of leading) {
      const past = transformBatch(lead, [carriedMine], { order: "before" });
      const next = transformBatch(carriedMine, [lead]);
      const afterMine = applyBatch(cursor, carriedMine.ops);
      const afterLead = applyBatch(cursor, lead.ops);
      if (
        past.isErr() ||
        next.isErr() ||
        afterMine.isErr() ||
        afterLead.isErr()
      )
        break;
      const pastApplied = applyBatch(afterMine.value.document, past.value.ops);
      const nextApplied = applyBatch(afterLead.value.document, next.value.ops);
      if (pastApplied.isErr() || nextApplied.isErr()) break;
      carried.push({
        ...past.value,
        revision: lead.revision,
        effects: pastApplied.value.effects,
      });
      carriedMine = {
        ...next.value,
        revision: headRev + 1,
        effects: nextApplied.value.effects,
      };
      cursor = afterLead.value.document;
    }
    if (carried.length === leading.length)
      return Result.ok({ batch: moved.value, leading: carried });
    // Later batches lose this one, cross `leading`, and take it back as
    // rebased: one that depends on it cannot lose it and is dropped.
    const made = applyBatch(at, entry.batch.ops);
    if (made.isErr())
      return panic(
        "Pending batch must replay in its original coordinate space.",
      );
    const undone = applyBatch(made.value.document, entry.inverse);
    const base = documentAfter(at, leading);
    const redone =
      base === undefined ? undefined : applyBatch(base, moved.value.ops);
    if (undone.isErr() || redone === undefined || redone.isErr())
      return panic("A rebased pending batch must apply where it rebased.");
    return Result.ok({
      batch: moved.value,
      leading: [
        {
          ...entry.batch,
          ops: entry.inverse,
          revision: headRev + 1,
          effects: undone.value.effects,
        },
        ...leading,
        {
          ...moved.value,
          revision: headRev + 1,
          effects: redone.value.effects,
        },
      ],
    });
  };

  /**
   * Rebase pending batches, made one after another from `original`, over
   * `leading`. A batch that does not rebase is dropped (`forced` drops the
   * first one whatever it is, as a rejection does), and the batches after
   * it rebase as if it had never been made: over its inverse first. Only a
   * batch that depends on it, editing what it put there, fails that and is
   * dropped too. Returns the kept batches and `leading` carried past them.
   */
  const rebaseRun = ({
    run,
    original,
    leading,
    forced,
  }: {
    run: readonly PendingEntry[];
    original: Document;
    leading: readonly SequencedBatch[];
    forced?: BatchRejection;
  }): { kept: PendingEntry[]; leading: SequencedBatch[] } => {
    const kept: PendingEntry[] = [];
    let cursor = original;
    let carried = [...leading];
    for (const [index, entry] of run.entries()) {
      const made = applyBatch(cursor, entry.batch.ops);
      if (made.isErr())
        return panic(
          "Pending batch must replay in its original coordinate space.",
        );
      const passed =
        index === 0 && forced !== undefined
          ? Result.err(forced)
          : passOver(entry, carried, cursor);
      if (passed.isOk()) {
        kept.push({ ...entry, batch: passed.value.batch });
        carried = passed.value.leading;
      } else {
        drop(entry, passed.error);
        const undone = applyBatch(made.value.document, entry.inverse);
        if (undone.isErr())
          return panic("Recorded inverse must apply after its batch.");
        carried = [
          {
            ...entry.batch,
            ops: entry.inverse,
            revision: headRev + 1,
            effects: undone.value.effects,
          },
          ...carried,
        ];
      }
      cursor = made.value.document;
    }
    return { kept, leading: carried };
  };

  // Batches applied in turn as one, under the identity of `source`.
  const joined = (
    batches: readonly SequencedBatch[],
    source: Omit<SequencedBatch, "effects">,
  ): SequencedBatch => {
    const only = batches.at(0);
    if (batches.length === 1 && only !== undefined)
      return {
        ...only,
        opId: source.opId,
        actor: source.actor,
        revision: source.revision,
      };
    return {
      ...source,
      ops: batches.flatMap(({ ops }) => ops),
      effects: batches.flatMap(({ effects }) => effects ?? []),
    };
  };

  const replay = (): void => {
    current = confirmed;
    const surviving: PendingEntry[] = [];
    for (const [index, entry] of entries.entries()) {
      const applied = applyBatch(current, entry.batch.ops);
      if (applied.isErr()) {
        drop(
          entry,
          new BatchRejection({
            reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
            message: applied.error.message,
          }),
        );
        for (const rest of entries.slice(index + 1)) {
          drop(
            rest,
            new BatchRejection({
              reason: BATCH_REJECTION_REASONS.CONFLICT,
              message:
                "An earlier optimistic dependency could not be replayed.",
            }),
          );
        }
        break;
      }
      current = applied.value.document;
      surviving.push({
        ...entry,
        inverse: applied.value.inverse,
        effects: applied.value.effects,
      });
    }
    entries = surviving;
  };

  const enqueue = (
    batch: DocumentBatch,
  ): Result<DocumentBatch, BatchRejection> => {
    const validated = validateDocumentBatch(batch);
    if (validated.isErr()) {
      return validated;
    }
    batch = validated.value;
    if (
      batch.revision !== undefined ||
      batch.baseRev !== headRev ||
      localIds.has(batch.opId)
    ) {
      return Result.err(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
          message:
            "Local batch must name the current revision and a fresh operation id.",
        }),
      );
    }
    const applied = applyBatch(current, batch.ops);
    if (applied.isErr()) {
      return Result.err(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
          message: applied.error.message,
        }),
      );
    }
    const normalizedOps: DocumentOp[] = [];
    for (const [index, op] of validated.value.ops.entries()) {
      const restored = restoreDocumentOp(op);
      if (restored.isErr()) {
        return Result.err(
          new BatchRejection({
            reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
            message: restored.error.message,
          }),
        );
      }
      const admitted = restored.value;
      const effect = applied.value.effects.at(index);
      normalizedOps.push(
        captureDocumentOp(
          admitted.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK &&
            effect?.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK
            ? { ...admitted, newHalf: effect.newHalf }
            : admitted,
        ),
      );
    }
    const normalized = { ...validated.value, ops: normalizedOps };
    entries.push({
      batch: normalized,
      inverse: applied.value.inverse,
      effects: applied.value.effects,
      status: PENDING_STATUSES.QUEUED,
    });
    localIds.add(batch.opId);
    current = applied.value.document;
    return Result.ok(normalized);
  };

  const nextSubmission = (): DocumentBatch | undefined => {
    if (inFlight !== undefined) {
      return undefined;
    }
    const first = entries.at(0);
    if (first === undefined) {
      return undefined;
    }
    const batch = { ...first.batch, baseRev: headRev };
    entries[0] = { ...first, batch, status: PENDING_STATUSES.SUBMITTED };
    inFlight = batch.opId;
    return batch;
  };

  const receiveAck = (ack: BatchAck): void => {
    entries = entries.map((entry) =>
      entry.batch.opId === ack.opId
        ? { ...entry, status: PENDING_STATUSES.ACKNOWLEDGED }
        : entry,
    );
    if (appliedIds.has(ack.opId) && inFlight === ack.opId) {
      inFlight = undefined;
    }
  };

  const receiveReject = (
    rejection: BatchReject,
  ): readonly ClientDocumentChange[] => {
    const index = entries.findIndex(
      ({ batch }) => batch.opId === rejection.opId,
    );
    if (index < 0) {
      if (inFlight === rejection.opId) {
        inFlight = undefined;
      }
      return [];
    }
    const entry = entries[index];
    if (entry === undefined) {
      return [];
    }
    const previous = current;
    const previousEntries = entries.slice();
    const noticesBefore = notices.length;
    const prefix = entries.slice(0, index);
    const original = applyBatch(
      confirmed,
      prefix.flatMap(({ batch }) => batch.ops),
    );
    if (original.isErr())
      return panic("Pending prefix must replay before a rejection is removed.");
    // Later batches were authored with the rejected edit present: they
    // rebase over its inverse, and only one that depends on it goes too.
    const rebased = rebaseRun({
      run: entries.slice(index),
      original: original.value.document,
      leading: [],
      forced: rejection.reason,
    });
    const inverse = joined(rebased.leading, {
      ...entry.batch,
      revision: headRev + 1,
    });
    const retained = [...prefix, ...rebased.kept];
    entries = retained;
    if (inFlight === rejection.opId) {
      inFlight = undefined;
    }
    replay();
    const source = {
      opId: entry.batch.opId,
      actor: entry.batch.actor,
      revision: headRev + 1,
    };
    if (notices.length > noticesBefore + 1)
      return [replacement(previous, previousEntries, [], source)];
    const landed = landing(previous, inverse);
    switch (landed) {
      case "exact":
        return [{ type: "rebased", batch: inverse }];
      case "offsets":
        return [replacement(previous, previousEntries, [], source, inverse)];
      case "elsewhere":
        return [replacement(previous, previousEntries, [], source)];
      default: {
        const exhaustive: never = landed;
        return exhaustive;
      }
    }
  };

  const receiveBroadcast = (
    batch: SequencedBatch,
  ): Result<readonly ClientDocumentChange[], BatchRejection> => {
    const checked = validateSequencedBatch(batch);
    if (checked.isErr()) return Result.err(checked.error);
    batch = checked.value;
    if (
      !Number.isSafeInteger(batch.revision) ||
      batch.revision < 1 ||
      batch.baseRev !== batch.revision - 1
    ) {
      return Result.err(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
          message: "Broadcast does not name an adjacent journal revision.",
        }),
      );
    }
    const existing =
      history.get(batch.revision) ?? received.get(batch.revision);
    if (existing !== undefined) {
      return equalForStaleness(existing, batch)
        ? Result.ok([])
        : Result.err(
            new BatchRejection({
              reason: BATCH_REJECTION_REASONS.CONFLICT,
              message:
                "Broadcast conflicts with a previously delivered revision.",
            }),
          );
    }
    received.set(batch.revision, batch);
    const changes: ClientDocumentChange[] = [];
    while (received.has(headRev + 1)) {
      const next = received.get(headRev + 1);
      if (next === undefined) {
        break;
      }
      if (appliedIds.has(next.opId)) {
        return Result.err(
          new BatchRejection({
            reason: BATCH_REJECTION_REASONS.CONFLICT,
            message: "An operation id appears at two journal revisions.",
          }),
        );
      }
      const applied = applyBatch(confirmed, next.ops);
      if (applied.isErr()) {
        return Result.err(
          new BatchRejection({
            reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
            message: applied.error.message,
          }),
        );
      }
      const local = entries.find(
        ({ batch: pendingBatch }) => pendingBatch.opId === next.opId,
      );
      if (local !== undefined && local.batch.actor !== next.actor) {
        return Result.err(
          new BatchRejection({
            reason: BATCH_REJECTION_REASONS.CONFLICT,
            message:
              "A broadcast reuses a pending operation id under another actor.",
          }),
        );
      }
      const previous = current;
      const previousEntries = entries.slice();
      const noticesBefore = notices.length;
      let rebasedRemote: SequencedBatch | undefined;
      if (localIds.has(next.opId)) {
        entries = entries.filter(
          ({ batch: pendingBatch }) => pendingBatch.opId !== next.opId,
        );
        if (inFlight === next.opId) {
          inFlight = undefined;
        }
      } else {
        const rebased = rebaseRun({
          run: entries,
          original: confirmed,
          leading: [next],
        });
        entries = rebased.kept;
        rebasedRemote = joined(rebased.leading, next);
      }
      confirmed = applied.value.document;
      headRev = next.revision;
      appliedIds.add(next.opId);
      received.delete(headRev);
      history.set(headRev, next);
      replay();
      const source = {
        opId: next.opId,
        actor: next.actor,
        revision: next.revision,
      };
      // An own batch coming back moves nothing the client had not shown.
      if (notices.length > noticesBefore)
        changes.push(replacement(previous, previousEntries, next.ops, source));
      else if (rebasedRemote !== undefined) {
        const landed =
          previousEntries.length > 0
            ? landing(previous, rebasedRemote)
            : "exact";
        switch (landed) {
          case "exact":
            changes.push({ type: "rebased", batch: rebasedRemote });
            break;
          case "offsets":
            changes.push(
              replacement(
                previous,
                previousEntries,
                next.ops,
                source,
                rebasedRemote,
              ),
            );
            break;
          case "elsewhere":
            changes.push(
              replacement(previous, previousEntries, next.ops, source),
            );
            break;
          default: {
            const exhaustive: never = landed;
            return exhaustive;
          }
        }
      }
    }
    return Result.ok(changes);
  };

  return {
    enqueue,
    nextSubmission,
    receiveAck,
    receiveReject,
    receiveBroadcast,
    get document(): Document {
      return current;
    },
    get headRev(): number {
      return headRev;
    },
    get pending(): readonly DocumentBatch[] {
      return entries
        .filter(({ status }) => status !== PENDING_STATUSES.ACKNOWLEDGED)
        .map(({ batch }) => batch);
    },
    get notices(): readonly ClientNotice[] {
      return notices.slice();
    },
  };
};
