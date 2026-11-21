import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  formatParaId,
  ID_KIND_LIMITS,
  ID_LEASE_KINDS,
  type IdLease,
  type LeasedId,
  OP_STORIES,
} from "@stll/docx-core/ops";

type ReplayRejection = "lease" | "suggestion" | "windowShape" | "identity";

type ReplayMember = {
  id: string;
  ops: readonly DocumentOp[];
  ids: readonly LeasedId[];
};

type AdmissionLease =
  | { type: "ordinary" }
  | { type: "bounded"; lease: IdLease; start: number }
  | { type: "fallback"; reason: ReplayRejection | "exhausted" };

/**
 * Once ordinary allocation succeeds, smaller paragraph IDs remain held
 * for the rest of a plain insertion window. Advance only when an actual
 * native insertion with exactly one new paragraph was admitted.
 */
export const startParagraphAdmissionCursor = (
  reason: ReplayRejection | undefined,
) => {
  let next: AdmissionLease =
    reason === undefined ? { type: "ordinary" } : { type: "fallback", reason };
  return {
    nextLease: (): AdmissionLease => next,
    commit: (paragraphs: readonly number[] | undefined): void => {
      if (next.type === "fallback") return;
      const id = paragraphs?.at(0);
      if (
        paragraphs?.length !== 1 ||
        id === undefined ||
        id < (next.type === "bounded" ? next.start : 1) ||
        id >= ID_KIND_LIMITS.paragraph
      ) {
        next = { type: "fallback", reason: "identity" };
        return;
      }
      const start = id + 1;
      next =
        start === ID_KIND_LIMITS.paragraph
          ? { type: "fallback", reason: "exhausted" }
          : {
              type: "bounded",
              start,
              lease: {
                blocks: [
                  {
                    kind: ID_LEASE_KINDS.PARAGRAPH,
                    start,
                    end: ID_KIND_LIMITS.paragraph,
                  },
                ],
              },
            };
    },
  };
};

/**
 * A plain insertion only adds its one paragraph id. Every smaller id that
 * blocked it at admission is still held by the package or the retired set,
 * or belongs to another admitted sibling. The member's own id remains free.
 * Keep this proof for the whole window, including replayed placements.
 */
export const planReplayParagraphIds = (
  rejection: ReplayRejection | undefined,
  members: readonly ReplayMember[],
):
  | { type: "fallback"; reason: ReplayRejection }
  | { type: "admitted"; leases: ReadonlyMap<string, IdLease> } => {
  if (rejection !== undefined) return { type: "fallback", reason: rejection };

  const leases = new Map<string, IdLease>();
  const seen = new Set<number>();
  for (const member of members) {
    const op = member.ops.at(0);
    const admitted = member.ids.at(0);
    if (
      member.ops.length !== 1 ||
      op?.type !== DOCUMENT_OP_TYPES.INSERT_BLOCKS ||
      op.story !== OP_STORIES.MAIN ||
      op.move === true ||
      (op.at.type !== "before" && op.at.type !== "after") ||
      op.blocks.length !== 1 ||
      member.ids.length !== 1 ||
      admitted?.kind !== ID_LEASE_KINDS.PARAGRAPH ||
      admitted.id < 1 ||
      admitted.id >= ID_KIND_LIMITS.paragraph ||
      !Number.isInteger(admitted.id) ||
      seen.has(admitted.id) ||
      op.blocks.at(0)?.paraId !== formatParaId(admitted.id)
    )
      return { type: "fallback", reason: "identity" };
    seen.add(admitted.id);
    leases.set(member.id, {
      blocks: [
        {
          kind: ID_LEASE_KINDS.PARAGRAPH,
          start: admitted.id,
          end: admitted.id + 1,
        },
      ],
    });
  }
  return { type: "admitted", leases };
};
