import { Result, TaggedError, type TaggedErrorClass } from "better-result";
import * as v from "valibot";

import {
  type BatchSubmission,
  createSequencer,
  DEFAULT_ID_BLOCK_SIZES,
  DOCUMENT_OP_SCHEMA_VERSION,
  type DocumentBatch,
  ID_LEASE_KINDS,
  type IdBlock,
  type IdLease,
  type IdLeaseGrant,
  type IdLeaseKind,
  type IdLeaseRefusal,
  type IdLeaseState,
  type SequencedBatch,
  validateSequencedBatch,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";

/**
 * What a room's authority keeps durably, in order: every sequenced batch and
 * every id block it leased. Replaying them over the room's base document
 * gives back the same sequencer: the same journal, the same lease table, and
 * the same answer to a batch submitted again under an operation id it
 * already sequenced.
 */
export type CanonicalRoomRecord =
  | { type: "sequenced"; batch: SequencedBatch }
  | { type: "leased"; grant: IdLeaseGrant };

const CanonicalRoomAuthorityErrorBase: TaggedErrorClass<"CanonicalRoomAuthorityError"> =
  TaggedError("CanonicalRoomAuthorityError");

/**
 * Why an authority could not be rebuilt from its records. `invalidRecord`: a
 * record is malformed. `invalidLeases`: the leased blocks do not form a lease
 * table. `replayDiverged`: a sequenced batch did not sequence again at its
 * own revision, so the records are not one room's journal.
 */
export class CanonicalRoomAuthorityError extends CanonicalRoomAuthorityErrorBase<{
  message: string;
  reason: "invalidRecord" | "invalidLeases" | "replayDiverged";
  cause?: unknown;
}> {}

/** What an authority call answers, and the records to store before anyone hears it. */
export type CanonicalRoomAuthorityResult<T> = {
  answer: T;
  records: readonly CanonicalRoomRecord[];
};

/**
 * A room's sequencer, rebuilt from its records. Each call returns the
 * records it added; the caller stores them durably before it broadcasts or
 * answers, so a successor rebuilt from the store never forgets an edit
 * somebody heard of.
 */
export type CanonicalRoomAuthority = {
  submit(batch: DocumentBatch): CanonicalRoomAuthorityResult<BatchSubmission>;
  grantIdLease(
    actor: string,
  ): CanonicalRoomAuthorityResult<Result<IdLease, IdLeaseRefusal>>;
  renewIdLease(
    actor: string,
    kind: IdLeaseKind,
  ): CanonicalRoomAuthorityResult<Result<IdBlock, IdLeaseRefusal>>;
  readonly document: Document;
  readonly headRev: number;
  readonly journal: readonly SequencedBatch[];
};

const GRANT = v.object({
  type: v.literal("leased"),
  grant: v.object({
    actor: v.pipe(v.string(), v.minLength(1)),
    block: v.object({
      kind: v.picklist(Object.values(ID_LEASE_KINDS)),
      start: v.number(),
      end: v.number(),
    }),
  }),
});
const SEQUENCED = v.object({
  type: v.literal("sequenced"),
  batch: v.unknown(),
});

/** A stored record, or undefined when it is not one. */
export const decodeRoomRecord = (
  value: unknown,
): CanonicalRoomRecord | undefined => {
  const grant = v.safeParse(GRANT, value);
  if (grant.success) return grant.output;
  const sequenced = v.safeParse(SEQUENCED, value);
  if (!sequenced.success) return undefined;
  const batch = validateSequencedBatch(sequenced.output.batch);
  return batch.isOk() ? { type: "sequenced", batch: batch.value } : undefined;
};

const authorityError = (
  reason: CanonicalRoomAuthorityError["reason"],
  message: string,
  cause?: unknown,
) =>
  Result.err(
    new CanonicalRoomAuthorityError({
      reason,
      message,
      ...(cause === undefined ? {} : { cause }),
    }),
  );

export type RestoreCanonicalRoomAuthorityOptions = {
  /** Revision 0: the document a canonical session reads from the room's DOCX. */
  document: Document;
  /** What earlier authorities of the room stored, in order. */
  records: readonly unknown[];
  /** How many ids one lease block holds, per kind; the same for every authority of a room. */
  idBlockSizes?: IdLeaseState["blockSizes"];
};

/**
 * Rebuild a room's authority from its records: every leased block is given
 * back to the lease table, then every sequenced batch is submitted again at
 * its own revision, in order, and must sequence there. Only then does the
 * authority take new submissions.
 */
export const restoreCanonicalRoomAuthority = ({
  document,
  records,
  idBlockSizes = DEFAULT_ID_BLOCK_SIZES,
}: RestoreCanonicalRoomAuthorityOptions): Result<
  CanonicalRoomAuthority,
  CanonicalRoomAuthorityError
> => {
  const grants: IdLeaseGrant[] = [];
  const journal: SequencedBatch[] = [];
  for (const value of records) {
    const record = decodeRoomRecord(value);
    if (record === undefined)
      return authorityError(
        "invalidRecord",
        "A stored room record is malformed.",
      );
    if (record.type === "leased") grants.push(record.grant);
    else journal.push(record.batch);
  }
  const created = Result.try(() =>
    createSequencer(document, {
      idLeases: { blockSizes: idBlockSizes, grants, minted: [] },
    }),
  );
  if (created.isErr())
    return authorityError(
      "invalidLeases",
      "The stored lease blocks do not form a lease table.",
      created.error,
    );
  const sequencer = created.value;
  for (const batch of journal) {
    const outcome = sequencer.submit({
      schema: DOCUMENT_OP_SCHEMA_VERSION,
      opId: batch.opId,
      actor: batch.actor,
      baseRev: batch.revision - 1,
      ops: batch.ops,
    });
    if (outcome.type !== "ack" || outcome.rev !== batch.revision)
      return authorityError(
        "replayDiverged",
        `The stored batch ${batch.opId} does not sequence again at revision ${batch.revision}.`,
        outcome.type === "reject" ? outcome.reason : undefined,
      );
  }

  const leasedSince = (before: number): CanonicalRoomRecord[] =>
    sequencer.idLeases.grants
      .slice(before)
      .map((grant) => ({ type: "leased", grant }));

  return Result.ok({
    submit: (batch) => {
      const before = sequencer.headRev;
      const answer = sequencer.submit(batch);
      const sequenced =
        sequencer.headRev > before ? sequencer.broadcasts.at(-1) : undefined;
      return {
        answer,
        records:
          sequenced === undefined
            ? []
            : [{ type: "sequenced", batch: sequenced }],
      };
    },
    grantIdLease: (actor) => {
      const before = sequencer.idLeases.grants.length;
      const answer = sequencer.grantIdLease(actor);
      return { answer, records: leasedSince(before) };
    },
    renewIdLease: (actor, kind) => {
      const before = sequencer.idLeases.grants.length;
      const answer = sequencer.renewIdLease(actor, kind);
      return { answer, records: leasedSince(before) };
    },
    get document() {
      return sequencer.document;
    },
    get headRev() {
      return sequencer.headRev;
    },
    get journal() {
      return sequencer.broadcasts;
    },
  });
};
