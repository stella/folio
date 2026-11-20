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
