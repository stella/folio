import { DOCUMENT_OP_SCHEMA_VERSION } from "../../src/ops/types";
import type { DocumentBatch, SequencedBatch } from "../../src/ops/documentOps";

const SUBMISSION = {
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId: "batch-1",
  actor: "actor-1",
  baseRev: 0,
  ops: [],
} as const satisfies DocumentBatch;

// @ts-expect-error a submission has no assigned journal revision
const UNSEQUENCED: SequencedBatch = SUBMISSION;

const SEQUENCED = { ...SUBMISSION, revision: 1 } as const satisfies SequencedBatch;

// @ts-expect-error operation schema versions are an explicit wire contract
const WRONG_SCHEMA: DocumentBatch = { ...SUBMISSION, schema: 0 };

export type BatchProof = [typeof UNSEQUENCED, typeof SEQUENCED, typeof WRONG_SCHEMA];
