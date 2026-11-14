/**
 * Paragraph restorations against concurrent edits of the same paragraph.
 *
 * A restoration (`replaceInline`, and each anchor paragraph of
 * `restoreCommentState`) states a paragraph's whole content before and after
 * it, but changes only one stretch of it: its span, the leaves between the
 * longest common prefix and the longest common suffix of the two contents,
 * where leaves compare with the records holding them. An edit lying outside
 * the span passes it: the edit's positions after the span move by the width
 * the restoration changes, and the restoration rebased over the edit states
 * the contents the edit leaves on either side. An edit reaching into the
 * span conflicts with it.
 */

import { Result } from "better-result";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { storyBody, storyParagraphs } from "../blocks";
import { structurallyEqual } from "../equality";
import { compareGaps, type Gap, type InlineNode, leafSpans } from "../leaves";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type OpStory,
  type TextPosition,
} from "../types";
import { BATCH_REJECTION_REASONS, BatchRejection } from "./envelope";
import { isInlineInsertion, sameParagraph } from "./points";

/**
 * Where a restoration changes a paragraph: from `from` up to `to` in the
 * content it expects, and up to `restoredTo` in the content it gives.
 */
export type RestoredSpan = { from: Gap; to: Gap; restoredTo: Gap };

type Unit = {
  /** A character of run text, or the whole leaf. */
  leaf: { char: string } | { node: InlineNode };
  /** The records holding the unit, outermost first; a character's text last. */
  ancestors: readonly InlineNode[];
  before: Gap;
  after: Gap;
};

const units = (content: readonly InlineNode[]): Unit[] =>
  leafSpans(content).flatMap(({ node, ancestors, before, after }): Unit[] => {
    if (node.type !== "text" || node.text.length <= 1)
      return [
        node.type === "text"
          ? {
              leaf: { char: node.text },
              ancestors: [...ancestors, node],
              before,
              after,
            }
          : { leaf: { node }, ancestors, before, after },
      ];
    // Offsets count UTF-16 code units, so a character here is one unit.
    const holders = [...ancestors, node];
    const { text } = node;
    return Array.from({ length: text.length }, (_, index): Unit => ({
      leaf: { char: text.charAt(index) },
      ancestors: holders,
      before:
        index === 0
          ? before
          : { offset: before.offset + index, zeroWidthBefore: 0 },
      after:
        index === text.length - 1
          ? after
          : { offset: before.offset + index + 1, zeroWidthBefore: 0 },
    }));
  });

/** A record's fields other than its children, and a text's other than its text. */
const ownFields = (node: InlineNode): Record<string, unknown> =>
  Object.fromEntries(
    Object.entries(node).filter(
      ([key]) =>
        key !== (node.type === "hyperlink" ? "children" : "content") &&
        !(node.type === "text" && key === "text"),
    ),
  );

/**
 * Whether two units are the same leaf held by alike records, with the same
 * record boundaries against their neighbours on one side.
 */
const sameUnit = (
  left: Unit,
  right: Unit,
  leftNeighbour: Unit | undefined,
  rightNeighbour: Unit | undefined,
): boolean =>
  ("char" in left.leaf
    ? "char" in right.leaf && left.leaf.char === right.leaf.char
    : "node" in right.leaf &&
      structurallyEqual(left.leaf.node, right.leaf.node)) &&
  left.ancestors.length === right.ancestors.length &&
  left.ancestors.every((record, level) => {
    const other = right.ancestors[level];
    return (
      other !== undefined &&
      record.type === other.type &&
      structurallyEqual(ownFields(record), ownFields(other)) &&
      (leftNeighbour?.ancestors[level] === record) ===
        (rightNeighbour?.ancestors[level] === other)
    );
  });

const START: Gap = { offset: 0, zeroWidthBefore: 0 };

/** The gap before unit `index`, or after the last unit. */
const gapAt = (list: readonly Unit[], index: number): Gap =>
  list[index]?.before ?? list.at(-1)?.after ?? START;

/**
 * The span a restoration of `expected` to `content` changes: the stretch
 * between their longest common prefix and longest common suffix of units,
 * the suffix taking only what the prefix leaves. A change in a record's own
 * fields alone spans that record's leaves; a change of record boundaries
 * alone is an empty span at the boundary.
 */
export const restoredSpan = (
  expected: readonly ParagraphContent[],
  content: readonly ParagraphContent[],
): RestoredSpan => {
  const before = units(expected);
  const after = units(content);
  let prefix = 0;
  while (
    prefix < before.length &&
    prefix < after.length &&
    sameUnit(
      before[prefix] as Unit,
      after[prefix] as Unit,
      before[prefix - 1],
      after[prefix - 1],
    )
  )
    prefix += 1;
  const room = Math.min(before.length, after.length) - prefix;
  let suffix = 0;
  while (
    suffix < room &&
    sameUnit(
      before[before.length - 1 - suffix] as Unit,
      after[after.length - 1 - suffix] as Unit,
      before[before.length - suffix],
      after[after.length - suffix],
    )
  )
    suffix += 1;
  return {
    from: gapAt(before, prefix),
    to: gapAt(before, before.length - suffix),
    restoredTo: gapAt(after, after.length - suffix),
  };
};

const gapOf = ({ offset, zeroWidthBefore }: TextPosition): Gap => ({
  offset,
  zeroWidthBefore: zeroWidthBefore ?? 0,
});

/**
 * A position of the expected content in the restored one: unchanged up to
 * the span's start, moved by the span's change of width from its end on, and
 * undefined strictly inside it.
 */
const pastSpan = (
  point: TextPosition,
  span: RestoredSpan,
): TextPosition | undefined => {
  const gap = gapOf(point);
  if (compareGaps(gap, span.from) <= 0) return point;
  if (compareGaps(gap, span.to) < 0) return undefined;
  if (gap.offset === span.to.offset)
    return {
      ...point,
      offset: span.restoredTo.offset,
      zeroWidthBefore:
        gap.zeroWidthBefore -
        span.to.zeroWidthBefore +
        span.restoredTo.zeroWidthBefore,
    };
  return {
    ...point,
    offset: point.offset + span.restoredTo.offset - span.to.offset,
  };
};

/** Whether a range's leaves reach into the span's. */
const reaches = (
  from: TextPosition,
  to: TextPosition,
  span: RestoredSpan,
): boolean =>
  compareGaps(gapOf(from), span.to) < 0 &&
  compareGaps(span.from, gapOf(to)) < 0;

/** The edits a restoration passes: they edit one paragraph by position. */
type ParagraphEdit = Extract<
  DocumentOp,
  {
    type:
      | typeof DOCUMENT_OP_TYPES.INSERT_TEXT
      | typeof DOCUMENT_OP_TYPES.INSERT_CONTENT
      | typeof DOCUMENT_OP_TYPES.SPLIT_INLINE
      | typeof DOCUMENT_OP_TYPES.JOIN_INLINE
      | typeof DOCUMENT_OP_TYPES.DELETE_RANGE
      | typeof DOCUMENT_OP_TYPES.SET_RUN_PROPS;
  }
>;

const paragraphEdit = (op: DocumentOp): ParagraphEdit | undefined => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
      return op;
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
      return isInlineInsertion(op) ? op : undefined;
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return sameParagraph(op.from, op.to) ? op : undefined;
    default:
      return undefined;
  }
};

const editedParagraph = (edit: ParagraphEdit): TextPosition =>
  "at" in edit ? edit.at : edit.from;

/** The edit's positions moved past the span; undefined when it reaches into it. */
const editPast = (
  edit: ParagraphEdit,
  span: RestoredSpan,
): ParagraphEdit | undefined => {
  if ("at" in edit) {
    const at = pastSpan(edit.at, span);
    return at === undefined ? undefined : { ...edit, at };
  }
  if (reaches(edit.from, edit.to, span)) return undefined;
  const from = pastSpan(edit.from, span);
  const to = pastSpan(edit.to, span);
  return from === undefined || to === undefined
    ? undefined
    : { ...edit, from, to };
};

/** A paragraph a restoration rewrites: what it expects and what it gives. */
type Rewrite = {
  story: OpStory;
  blockId: string;
  expected: readonly ParagraphContent[];
  content: readonly ParagraphContent[];
};

const rewriteOf = (op: DocumentOp, at: TextPosition): Rewrite | undefined => {
  if (op.type === DOCUMENT_OP_TYPES.REPLACE_INLINE)
    return sameParagraph(op, at)
      ? {
          story: op.story,
          blockId: op.blockId,
          expected: op.expected,
          content: op.content,
        }
      : undefined;
  if (op.type !== DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE) return undefined;
  const expected = op.expected.anchors.filter((anchor) =>
    sameParagraph(anchor, at),
  );
  const state = op.state.anchors.filter((anchor) => sameParagraph(anchor, at));
  const [before] = expected;
  const [after] = state;
  return before === undefined ||
    after === undefined ||
    expected.length !== 1 ||
    state.length !== 1
    ? undefined
    : {
        story: before.story,
        blockId: before.blockId,
        expected: before.content,
        content: after.content,
      };
};

const inMainStory = (edit: ParagraphEdit): ParagraphEdit => {
  const main = (position: TextPosition): TextPosition => ({
    ...position,
    story: OP_STORIES.MAIN,
  });
  return "at" in edit
    ? { ...edit, at: main(edit.at) }
    : { ...edit, from: main(edit.from), to: main(edit.to) };
};

/**
 * A paragraph's content after an edit, applied to that paragraph alone: the
 * edit reads nothing outside its paragraph, wherever the paragraph is.
 */
const editedContent = (
  edit: ParagraphEdit,
  blockId: string,
  content: readonly ParagraphContent[],
): readonly ParagraphContent[] | undefined => {
  const document: Document = {
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: blockId, content: [...content] },
        ],
      },
    },
  };
  const applied = applyDocumentOps(document, [inMainStory(edit)]);
  if (applied.isErr()) return undefined;
  return storyParagraphs(
    storyBody(applied.value.document, OP_STORIES.MAIN),
  ).find(({ paragraph }) => paragraph.paraId === blockId)?.paragraph.content;
};

const withRewrite = (
  op: DocumentOp,
  rewrite: Rewrite,
  expected: readonly ParagraphContent[],
  content: readonly ParagraphContent[],
): DocumentOp => {
  if (op.type === DOCUMENT_OP_TYPES.REPLACE_INLINE)
    return { ...op, expected, content };
  if (op.type !== DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE) return op;
  const replaced = (
    anchors: typeof op.expected.anchors,
    paragraph: readonly ParagraphContent[],
  ) =>
    anchors.map((anchor) =>
      sameParagraph(anchor, rewrite)
        ? { ...anchor, content: paragraph }
        : anchor,
    );
  return {
    ...op,
    expected: {
      ...op.expected,
      anchors: replaced(op.expected.anchors, expected),
    },
    state: { ...op.state, anchors: replaced(op.state.anchors, content) },
  };
};

const spanConflict = (op: DocumentOp, over: DocumentOp) =>
  Result.err(
    new BatchRejection({
      reason: BATCH_REJECTION_REASONS.BLOCK_CONFLICT,
      message: "A concurrent edit reached the content a restoration changes.",
      opType: op.type,
      overType: over.type,
    }),
  );

/**
 * A restoration and a positional edit of one of its paragraphs. Undefined
 * for every other pair, and for one whose edit does not apply to the
 * paragraph alone, which the block footprint then decides.
 */
export const transformRestorePair = (
  op: DocumentOp,
  over: DocumentOp,
): Result<readonly DocumentOp[], BatchRejection> | undefined => {
  const overEdit = paragraphEdit(over);
  const opEdit = paragraphEdit(op);
  const edit = overEdit ?? opEdit;
  if (edit === undefined) return undefined;
  const restoration = overEdit === undefined ? over : op;
  const rewrite = rewriteOf(restoration, editedParagraph(edit));
  if (rewrite === undefined) return undefined;
  const span = restoredSpan(rewrite.expected, rewrite.content);
  const moved = editPast(edit, span);
  if (moved === undefined) return spanConflict(op, over);
  // An edit rebased over a restoration moves past its span.
  if (opEdit !== undefined && overEdit === undefined) return Result.ok([moved]);
  // A restoration rebased over an edit states what the edit leaves.
  const expected = editedContent(edit, rewrite.blockId, rewrite.expected);
  const content = editedContent(moved, rewrite.blockId, rewrite.content);
  if (expected === undefined || content === undefined) return undefined;
  return Result.ok([withRewrite(op, rewrite, expected, content)]);
};
