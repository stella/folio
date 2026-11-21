import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  formatParaId,
  locateParagraphThroughTextBoxes,
  type OpStory,
  paragraphIdsIn as modelParagraphIdsIn,
  paragraphsThroughTextBoxes,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";

/** A paragraph-side insertion keeps its anchor and introduces only these paragraph IDs. */
const paragraphIdsInInsertionOrder = (ops: readonly DocumentOp[]) => {
  const [op] = ops;
  if (
    ops.length !== 1 ||
    op?.type !== DOCUMENT_OP_TYPES.INSERT_BLOCKS ||
    (op.at.type !== "before" && op.at.type !== "after") ||
    op.move === true
  )
    return undefined;
  const ids: number[] = [];
  for (const { paraId } of op.blocks) {
    if (paraId === undefined) return undefined;
    const id = Number.parseInt(paraId, 16);
    if (!Number.isSafeInteger(id) || formatParaId(id) !== paraId)
      return undefined;
    ids.push(id);
  }
  if (modelParagraphIdsIn(op.blocks).length !== ids.length) return undefined;
  return ids;
};

export const plainInsertedParagraphIds = (ops: readonly DocumentOp[]) =>
  paragraphIdsInInsertionOrder(ops)?.sort((left, right) => left - right);

type InsertedCommentRangeOptions = {
  before: Document;
  after: Document;
  story: OpStory;
  ops: readonly DocumentOp[];
};

/** A compiled paragraph insertion introduces its fresh IDs in physical order. */
export const insertedCommentRange = ({
  before,
  after,
  story,
  ops,
}: InsertedCommentRangeOptions) => {
  const op = ops.at(0);
  if (
    op?.type === DOCUMENT_OP_TYPES.INSERT_BLOCKS &&
    op.story === story &&
    paragraphIdsInInsertionOrder(ops) !== undefined
  ) {
    const firstId = op.blocks.at(0)?.paraId;
    const first =
      firstId === undefined
        ? undefined
        : locateParagraphThroughTextBoxes(after, story, firstId);
    if (first !== undefined) {
      const inserted = op.blocks;
      const last = first.blocks[first.index + inserted.length - 1];
      if (
        last?.type === "paragraph" &&
        inserted.every(({ paraId }, offset) => {
          const actual = first.blocks[first.index + offset];
          return actual?.type === "paragraph" && actual.paraId === paraId;
        })
      )
        return { firstAdded: first.paragraph, lastAdded: last };
    }
  }
  // A compound or structural insertion can introduce paragraphs in several
  // block lists, including text boxes. Their story order defines the range.
  const known = new Set(
    paragraphsThroughTextBoxes(before, story).map(
      ({ paragraph }) => paragraph.paraId,
    ),
  );
  const added = paragraphsThroughTextBoxes(after, story).filter(
    ({ paragraph: { paraId } }) => paraId !== undefined && !known.has(paraId),
  );
  return {
    firstAdded: added.at(0)?.paragraph,
    lastAdded: added.at(-1)?.paragraph,
  };
};
