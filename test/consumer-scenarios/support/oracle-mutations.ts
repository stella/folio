import { RELATIONS, type Relation } from "./relation-contract.ts";

type Probe = { bug: string; assertion: string };

export const RELATION_MUTATIONS = {
  directTracked: {
    bug: "Tracked replacement writes a different payload than direct replacement.",
    assertion: "[directTracked]",
  },
  rejectAll: { bug: "Rejecting tracked revisions accepts them instead.", assertion: "[rejectAll]" },
  saveIdempotent: {
    bug: "Saving a reopened package changes its document part.",
    assertion: "[saveIdempotent]",
  },
  undo: { bug: "Undo reports success but leaves a new paragraph behind.", assertion: "[undo]" },
  batchSequential: {
    bug: "A multi-operation batch adds an unrequested paragraph.",
    assertion: "[batchSequential]",
  },
  readerStability: {
    bug: "The reopened content reader returns corrupted text.",
    assertion: "[readerStability]",
  },
} as const satisfies Record<Relation, Probe>;

export const ORACLE_MUTATIONS = {
  ...RELATION_MUTATIONS,
  requestedOutcome: {
    bug: "An applied replacement writes a different requested payload.",
    assertion: "the result is not what was asked",
  },
  saveRoundtrip: {
    bug: "Saving persists a different document than the live reviewer.",
    assertion: "the reopened package shows something else",
  },
  readerAgreement: {
    bug: "The content reader disagrees with the snapshot reader.",
    assertion: "snapshot vs getContent()",
  },
} as const;
export type OracleMutation = keyof typeof ORACLE_MUTATIONS;

export const relationForOracle = (oracle: OracleMutation): Relation | null =>
  RELATIONS.find((relation) => relation === oracle) ?? null;
