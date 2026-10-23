import { type Relation, RELATIONS } from "./relation-contract.js";

/**
 * Stable issue identities for every consumer-oracle and metamorphic check.
 * Keep fixture names, modes, seeds, and rendered values out of these keys.
 */
export const FAILURE_INVARIANTS = {
  requestedOutcome: { title: "Requested outcome matches the document" },
  numberingKind: { title: "Numbering kind is preserved" },
  numberingLabel: { title: "Numbering label is preserved" },
  numberingLevel: { title: "Numbering level is preserved" },
  commentLifetime: { title: "Comments do not outlive their anchors" },
  commentAnchor: { title: "Comment anchors are preserved" },
  blockPlacement: { title: "Blocks retain their requested placement" },
  rejectionRoundTrip: { title: "Rejecting changes restores the document" },
  directTracked: { title: "Accepted tracked edits equal direct edits" },
  rejectAll: { title: "Rejecting tracked edits restores the fixture" },
  saveIdempotent: { title: "Saving a reopened document is idempotent" },
  undo: { title: "Undo restores the pre-batch document" },
  batchSequential: { title: "Batch and sequential edits are equivalent" },
  readerStability: { title: "Readers are stable after save and reopen" },
} as const;

export type FailureInvariant = keyof typeof FAILURE_INVARIANTS;

/** Complete relation-to-invariant ownership; adding a relation requires a decision here. */
export const RELATION_INVARIANTS = {
  directTracked: "directTracked",
  rejectAll: "rejectAll",
  saveIdempotent: "saveIdempotent",
  undo: "undo",
  batchSequential: "batchSequential",
  readerStability: "readerStability",
} as const satisfies Record<Relation, FailureInvariant>;

const RULES: readonly { invariant: FailureInvariant; pattern: RegExp }[] = [
  ...RELATIONS.map((relation) => ({
    invariant: RELATION_INVARIANTS[relation],
    pattern: new RegExp(`\\[${relation}\\]`, "u"),
  })),
  {
    invariant: "rejectionRoundTrip",
    pattern: /rejecting every change|reject-all document differs/iu,
  },
  { invariant: "commentLifetime", pattern: /comment .* outlived/iu },
  { invariant: "commentAnchor", pattern: /no comment|comment .*anchor/iu },
  { invariant: "numberingLevel", pattern: /listLevel is|numbering level/iu },
  {
    invariant: "numberingLabel",
    pattern: /numbering format|label is|displayed label/iu,
  },
  { invariant: "numberingKind", pattern: /kind is .*expected/iu },
  {
    invariant: "blockPlacement",
    pattern: /block .*placement|block texts differ|block order/iu,
  },
  { invariant: "requestedOutcome", pattern: /result is not what was asked/iu },
];

export const invariantOfFailure = (
  assertion: string,
): FailureInvariant | null =>
  RULES.find(({ pattern }) => pattern.test(assertion))?.invariant ?? null;
