import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type EditorIntent,
  formatParaId,
  ID_KIND_LIMITS,
  ID_LEASE_KINDS,
  type IdLease,
  type LeasedId,
  OP_STORIES,
} from "@stll/docx-core/ops";

import type {
  FolioDocumentOperation,
  FolioDocumentOperationMode,
} from "../document-operations";

type ReplayOperation = {
  type: string;
  operation: FolioDocumentOperation;
  intents: readonly EditorIntent[];
};

type ReplayMember = {
  member: { ops: readonly DocumentOp[]; ids: readonly LeasedId[] };
  operation: ReplayOperation;
};

type ReplayParagraphIds =
  | {
      type: "fallback";
      reason: "lease" | "suggestion" | "windowShape" | "identity";
    }
  | { type: "admitted"; leases: ReadonlyMap<string, IdLease> };

const isPlainInsertion = ({
  type,
  operation,
  intents,
}: ReplayOperation): boolean => {
  const intent = intents.at(0);
  return (
    type === "plain" &&
    (operation.type === "insertBeforeBlock" ||
      operation.type === "insertAfterBlock") &&
    operation.moveId === undefined &&
    intents.length === 1 &&
    intent?.type === "insertParagraphs" &&
    intent.story === OP_STORIES.MAIN &&
    (intent.at.type === "before" || intent.at.type === "after") &&
    intent.move === undefined &&
    intent.newList === undefined &&
    intent.paragraphs.length === 1 &&
    intent.paragraphs.every(
      (paragraph) =>
        paragraph.sectionProperties === undefined &&
        paragraph.content.every(
          (node) =>
            node.type === "run" &&
            node.content.every(({ type: leaf }) => leaf === "text"),
        ),
    )
  );
};

type ReplayParagraphIdsOptions = {
  sessionLease: IdLease | undefined;
  mode: FolioDocumentOperationMode | undefined;
  window: readonly ReplayOperation[];
  ordered: readonly ReplayMember[];
};

/**
 * A plain insertion only adds its one paragraph id. Every smaller id that
 * blocked it at admission is still held by the package or the retired set,
 * or belongs to another admitted sibling. The member's own id remains free.
 * Keep this proof for the whole window, including replayed placements.
 */
export const planReplayParagraphIds = ({
  sessionLease,
  mode,
  window,
  ordered,
}: ReplayParagraphIdsOptions): ReplayParagraphIds => {
  if (sessionLease !== undefined) return { type: "fallback", reason: "lease" };
  if (mode === "suggested") return { type: "fallback", reason: "suggestion" };
  if (
    !window.every(isPlainInsertion) ||
    !ordered.every(({ operation }) => isPlainInsertion(operation))
  )
    return { type: "fallback", reason: "windowShape" };

  const leases = new Map<string, IdLease>();
  const seen = new Set<number>();
  for (const { member, operation } of ordered) {
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
    leases.set(operation.operation.id, {
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
