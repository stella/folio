import type { createTrackedPlan } from "../../src/ops/plan";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../../src/ops/types";

declare const plan: ReturnType<typeof createTrackedPlan>;

const join = {
  type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
  story: OP_STORIES.MAIN,
  blockId: "10000001",
  nextBlockId: "10000002",
} as const;

plan.join(join);
// @ts-expect-error A tracked join must pass through property restoration.
plan.append(join);
// @ts-expect-error The plan owns whether a join carries a revision.
plan.join({ ...join, revision: { id: 1, author: "A", date: "2026-01-01" } });
// @ts-expect-error A tracked join keeps the trailing paragraph's mark.
plan.join({ ...join, survivor: "first" });
