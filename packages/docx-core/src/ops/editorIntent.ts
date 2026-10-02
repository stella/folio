/** One editor intent, compiled to direct or tracked document operations. */
import { Result, panic } from "better-result";

import {
  MAX_REVISION_ID,
  type Document,
  type Paragraph,
  type ParagraphContent,
  type TextFormatting,
} from "../model/document";
import { sameBlockList, storyBody, storyParagraphs } from "./blocks";
import { IDENTITY_SPACES, idKey, packageIdentityKeys, packageParagraphIds } from "./ids";
import { leafSpans, zeroWidthLeavesAt } from "./leaves";
import { paragraphLength, paragraphLogicalText } from "./offsets";
import { appendTrackedDeletion, createTrackedPlan, selectedParagraphRuns } from "./plan";
import { planTrackedReplace, rangeStartAfterDeletion } from "./rangeReplacement";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { structurallyEqual } from "./equality";
import { locateTableRow, tableRowAnchor } from "./tableLocation";
import { tableGrid } from "./tableGrid";
import { isRemovedRevisionNode, paragraphPropertiesOf } from "./review";
import {
  DOCUMENT_OP_TYPES,
  SECTION_BOUNDARY_POLICIES,
  PROPERTY_REVIEW_POLICIES,
  EMPTY_PROPERTY_SETS,
  SPLIT_HALVES,
  type DocumentOp,
  type TableEditOp,
  type NewIds,
  type OpStory,
  type RevisionStamp,
  type TextPosition,
} from "./types";

/** Mode metadata is supplied once, independently of the semantic table intent. */
export type TableIntentOperation = {
  [Kind in TableEditOp["type"]]: Omit<Extract<TableEditOp, { type: Kind }>, "revision" | "newIds">;
}[TableEditOp["type"]];

/** Positions use canonical physical offsets, including retained deleted content. */
export type EditorIntent =
  | { type: "table"; operation: TableIntentOperation }
  | { type: "replaceText"; from: TextPosition; to: TextPosition; text: string }
  | { type: "splitParagraph"; at: TextPosition; to?: TextPosition; newBlockId: string }
  | { type: "joinParagraphs"; story: OpStory; blockId: string; nextBlockId: string };

/** Revision metadata and every fresh identity are supplied before compilation. */
export type EditorIntentMode =
  | { type: "editing"; newIds?: NewIds }
  | { type: "suggesting"; revision: RevisionStamp; newIds: NewIds };

export type CompileEditorIntentOptions = { intent: EditorIntent; mode: EditorIntentMode };
export type CompiledEditorIntent = { ops: DocumentOp[]; selection: TextPosition };

/** Fresh identities for one input, bounded by its paragraph leaves and ancestor records. */
export const allocateEditorIntentIds = (document: Document) => {
  const identities = packageIdentityKeys(document.package);
  const paragraphs = storyParagraphs(document.package.document);
  // Each leaf can start a deletion segment; each ancestor can be cut at both
  // endpoints. Paragraph joins need a mark and a property-change stamp.
  const demand =
    1 +
    paragraphs.reduce(
      (total, { paragraph }) =>
        total +
        3 +
        leafSpans(paragraph.content).reduce(
          (count, span) => count + 4 * (1 + span.ancestors.length),
          0,
        ),
      0,
    );
  const fresh = (space: string, count: number) => {
    const occupied = new Set(
      identities
        .filter((key) => key.startsWith(`${space}:`))
        .map((key) => Number(key.slice(space.length + 1))),
    );
    const ids: number[] = [];
    for (let id = 1; id <= MAX_REVISION_ID && ids.length < count; id += 1) {
      if (!occupied.has(id)) ids.push(id);
    }
    return ids;
  };
  const revision = fresh(IDENTITY_SPACES.REVISION, demand + 1);
  const revisionId = revision.at(0) ?? MAX_REVISION_ID + 1;
  const usedParagraphs = new Set(packageParagraphIds(document.package).map(idKey));
  let nextParagraph = 1;
  while (usedParagraphs.has(nextParagraph.toString(16).padStart(8, "0").toUpperCase()))
    nextParagraph += 1;
  return {
    revisionId,
    newBlockId: nextParagraph.toString(16).padStart(8, "0").toUpperCase(),
    newIds: { revision: revision.slice(1), control: fresh(IDENTITY_SPACES.CONTROL, demand) },
  };
};

/** Accepted-view text without changing the canonical paragraph or its offset space. */
export const paragraphVisibleText = (paragraph: Paragraph): string => {
  const text = paragraphLogicalText(paragraph);
  return leafSpans(paragraph.content)
    .filter(({ ancestors }) => !ancestors.some(isRemovedRevisionNode))
    .map(({ before, after }) => text.slice(before.offset, after.offset))
    .join("");
};

/**
 * Right-affine translation: a visible gap follows retained deleted content at
 * that gap, so a new insertion never lands inside a deletion wrapper.
 * Invalid offsets remain invalid for the operation boundary to diagnose.
 */
export const physicalOffsetAtVisibleOffset = (paragraph: Paragraph, offset: number): number => {
  if (!Number.isInteger(offset) || offset < 0) return -1;
  let visible = 0;
  for (const span of leafSpans(paragraph.content)) {
    if (span.ancestors.some(isRemovedRevisionNode)) continue;
    const width = span.after.offset - span.before.offset;
    if (offset < visible + width) return span.before.offset + offset - visible;
    visible += width;
  }
  return offset === visible ? paragraphLength(paragraph) : paragraphLength(paragraph) + 1;
};

/** Deleted paragraph marks concatenate physical paragraphs in the editing view. */
export const editorParagraphGroups = (document: Document, story: OpStory) => {
  const groups: { blockId: string; paragraphs: Paragraph[]; text: string }[] = [];
  let pending: Paragraph[] = [];
  let previous: ReturnType<typeof storyParagraphs>[number] | undefined;
  for (const location of storyParagraphs(storyBody(document, story))) {
    const { paragraph } = location;
    if (
      pending.length > 0 &&
      previous &&
      (!sameBlockList(previous.list, location.list) || location.index !== previous.index + 1)
    ) {
      groups.push({
        blockId: previous.paragraph.paraId ?? "",
        paragraphs: pending,
        text: pending.map(paragraphVisibleText).join(""),
      });
      pending = [];
    }
    previous = location;
    pending.push(paragraph);
    if (paragraph.pPrMark?.kind === "del" || paragraph.pPrMark?.kind === "moveFrom") continue;
    groups.push({
      blockId: paragraph.paraId ?? "",
      paragraphs: pending,
      text: pending.map(paragraphVisibleText).join(""),
    });
    pending = [];
  }
  if (pending.length > 0 && previous)
    groups.push({
      blockId: previous.paragraph.paraId ?? "",
      paragraphs: pending,
      text: pending.map(paragraphVisibleText).join(""),
    });
  return groups;
};

/** Translate an editing-view paragraph gap to its physical canonical address. */
export const physicalPositionAtEditorOffset = (
  document: Document,
  at: TextPosition,
): TextPosition => {
  const group = editorParagraphGroups(document, at.story).find(
    ({ blockId }) => idKey(blockId) === idKey(at.blockId),
  );
  if (group === undefined) return at;
  let remaining = at.offset;
  for (const [index, paragraph] of group.paragraphs.entries()) {
    const width = paragraphVisibleText(paragraph).length;
    if (remaining < width || index === group.paragraphs.length - 1) {
      return {
        story: at.story,
        blockId: paragraph.paraId ?? "",
        offset: physicalOffsetAtVisibleOffset(paragraph, remaining),
      };
    }
    remaining -= width;
  }
  return at;
};

const paragraphAt = (document: Document, at: TextPosition) =>
  storyParagraphs(storyBody(document, at.story)).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(at.blockId),
  )?.paragraph;

/** Capture authored formatting before deletion, using only visible runs. */
const formattingAt = (paragraph: Paragraph, offset: number): TextFormatting => {
  const spans = leafSpans(paragraph.content).filter(
    ({ ancestors }) => !ancestors.some(isRemovedRevisionNode),
  );
  const span =
    spans.find(({ before, after }) => before.offset <= offset && offset < after.offset) ??
    spans.findLast(({ after }) => after.offset <= offset) ??
    spans.find(({ before }) => before.offset > offset);
  return span?.ancestors.findLast((ancestor) => ancestor.type === "run")?.formatting ?? {};
};

const authoredFormatting = (
  document: Document,
  from: TextPosition,
  to: TextPosition,
): TextFormatting => {
  const paragraph = paragraphAt(document, from);
  if (paragraph === undefined)
    panic("A validated input paragraph must exist when reading authored formatting.");
  const collapsed = from.blockId === to.blockId && from.offset === to.offset;
  if (!collapsed) return formattingAt(paragraph, from.offset);
  const left = leafSpans(paragraph.content).some(
    ({ ancestors, before }) =>
      !ancestors.some(isRemovedRevisionNode) && before.offset < from.offset,
  );
  if (left) return formattingAt(paragraph, from.offset - 1);
  const group = editorParagraphGroups(document, from.story).find(({ paragraphs }) =>
    paragraphs.some(({ paraId }) => idKey(paraId ?? "") === idKey(from.blockId)),
  );
  const index =
    group?.paragraphs.findIndex(({ paraId }) => idKey(paraId ?? "") === idKey(from.blockId)) ?? 0;
  const previous = group?.paragraphs
    .slice(0, index)
    .findLast((item) => paragraphVisibleText(item) !== "");
  return previous === undefined
    ? formattingAt(paragraph, from.offset)
    : formattingAt(previous, paragraphLength(previous));
};

/** A new logical paragraph inherits the surviving paragraph mark's formatting. */
const splitParagraphFields = (document: Document, at: TextPosition) => {
  const survivor = editorParagraphGroups(document, at.story)
    .find(({ paragraphs }) =>
      paragraphs.some(({ paraId }) => idKey(paraId ?? "") === idKey(at.blockId)),
    )
    ?.paragraphs.at(-1);
  return survivor?.formatting === undefined ? {} : { formatting: survivor.formatting };
};

/** Compile one batch; the caller applies it atomically and journals its exact inverse. */
export const compileEditorIntent = (
  document: Document,
  { intent, mode }: CompileEditorIntentOptions,
): Result<CompiledEditorIntent, DocumentOpRefusal> => {
  const allocationFields = mode.newIds === undefined ? {} : { newIds: mode.newIds };
  const tracked =
    mode.type === "suggesting"
      ? { revision: mode.revision, newIds: mode.newIds }
      : allocationFields;
  let ops: DocumentOp[];
  let selection: TextPosition;
  switch (intent.type) {
    case "table": {
      const op = { ...intent.operation, ...tracked };
      const located = locateTableRow(document, op);
      if (located.isErr()) return Result.err(located.error);
      let blockId = op.blockId;
      if (op.type === DOCUMENT_OP_TYPES.DELETE_COLUMN) {
        const grid = tableGrid(located.value.table, op.type);
        if (grid.isErr()) return Result.err(grid.error);
        const rows = located.value.table.rows.map((row, index) => ({
          ...row,
          cells: row.cells.filter((_cell, cellIndex) => {
            const entry = grid.value.rows[index]?.find((cell) => cell.index === cellIndex);
            return (
              entry !== undefined &&
              (entry.start > op.column || entry.end <= op.column || entry.end - entry.start > 1)
            );
          }),
        }));
        const survivor =
          tableRowAnchor(rows) ??
          storyParagraphs(storyBody(document, op.story)).find((paragraph) =>
            sameBlockList(paragraph.list, located.value.list),
          )?.paragraph.paraId;
        if (survivor === undefined)
          return Result.err(
            new DocumentOpRefusal({
              opType: op.type,
              reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
              message: "Deleting the final column requires deleting the table.",
            }),
          );
        blockId = survivor;
      }
      ops = [op];
      selection = { story: op.story, blockId, offset: 0 };
      break;
    }
    case "replaceText": {
      const { from, to, text } = intent;
      const paragraph = paragraphAt(document, from);
      if (paragraph === undefined) {
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
            message: "The input paragraph does not exist.",
            opType: DOCUMENT_OP_TYPES.INSERT_TEXT,
          }),
        );
      }
      const runProps = authoredFormatting(document, from, to);
      // Both modes insert the same authored run; source XML attributes belong
      // to the existing run rather than to newly typed text.
      const content = (
        text === ""
          ? []
          : [{ type: "run", formatting: runProps, content: [{ type: "text", text }] }]
      ) satisfies ParagraphContent[];
      if (mode.type === "suggesting") {
        const planned = planTrackedReplace(document, {
          from,
          to,
          revision: mode.revision,
          newIds: mode.newIds,
          replacement: {
            paragraphs: [],
            tail: {
              openStart: 0,
              openEnd: 0,
              content,
            },
          },
        });
        if (planned.isErr()) return Result.err(planned.error);
        ops = planned.value;
        selection = { ...from, offset: from.offset + text.length };
        for (const op of ops) {
          if (
            op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS &&
            op.revision === undefined &&
            idKey(op.blockId) === idKey(selection.blockId)
          )
            selection = { ...selection, blockId: op.nextBlockId };
          if (op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT)
            selection = { ...op.at, offset: op.at.offset + text.length };
        }
        break;
      }
      ops = [];
      let survivorId = from.blockId;
      if (from.blockId !== to.blockId || from.offset !== to.offset) {
        const partitioned = selectedParagraphRuns(document, from, to);
        if (partitioned.isErr()) return Result.err(partitioned.error);
        survivorId = partitioned.value.at(0)?.at(-1)?.paragraph.paraId ?? from.blockId;
        for (const run of partitioned.value.toReversed()) {
          const last = run.at(-1)?.paragraph;
          if (!last) panic("A selected paragraph run must contain its survivor.");
          for (const { paragraph: item } of run.toReversed()) {
            const blockId = item.paraId ?? "";
            ops.push({
              type: DOCUMENT_OP_TYPES.DELETE_RANGE,
              ...allocationFields,
              from:
                idKey(blockId) === idKey(from.blockId)
                  ? from
                  : { story: from.story, blockId, offset: 0, zeroWidthBefore: 0 },
              to:
                idKey(blockId) === idKey(to.blockId)
                  ? to
                  : {
                      story: from.story,
                      blockId,
                      offset: paragraphLength(item),
                      zeroWidthBefore: zeroWidthLeavesAt(item.content, paragraphLength(item))
                        .length,
                    },
            });
          }
          for (const { paragraph: item } of run.slice(0, -1).toReversed()) {
            ops.push({
              type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
              story: from.story,
              blockId: item.paraId ?? "",
              nextBlockId: last.paraId ?? "",
              sectionBoundary: SECTION_BOUNDARY_POLICIES.REMOVE,
            });
          }
        }
      }
      const at = { ...from, blockId: survivorId };
      if (text !== "")
        ops.push({
          type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
          at,
          slice: { openStart: 0, openEnd: 0, content },
          ...allocationFields,
        });
      selection = { ...at, offset: from.offset + text.length };
      break;
    }
    case "splitParagraph": {
      if (intent.to === undefined) {
        ops = [
          {
            type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
            at: intent.at,
            newBlockId: intent.newBlockId,
            newHalf: SPLIT_HALVES.FIRST,
            newParagraph: splitParagraphFields(document, intent.at),
            ...tracked,
          },
        ];
        selection = { ...intent.at, offset: 0, zeroWidthBefore: 0 };
        break;
      }
      if (mode.type === "suggesting") {
        const plan = createTrackedPlan({ document, revision: mode.revision, newIds: mode.newIds });
        const deletion = appendTrackedDeletion({
          document,
          options: { from: intent.at, to: intent.to, revision: mode.revision, newIds: mode.newIds },
          plan,
        });
        if (deletion.isErr()) return Result.err(deletion.error);
        const at = rangeStartAfterDeletion({
          before: document,
          after: plan.document(),
          from: intent.at,
          to: intent.to,
        });
        const split = plan.append({
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          at,
          newBlockId: intent.newBlockId,
          newHalf: SPLIT_HALVES.FIRST,
          newParagraph: splitParagraphFields(plan.document(), at),
          newIds: mode.newIds,
          revision: mode.revision,
        });
        if (split.isErr()) return Result.err(split.error);
        ops = plan.ops;
        selection = { ...at, offset: 0, zeroWidthBefore: 0 };
        break;
      }
      const deletion = compileEditorIntent(document, {
        intent: { type: "replaceText", from: intent.at, to: intent.to, text: "" },
        mode,
      });
      if (deletion.isErr()) return deletion;
      ops = [
        ...deletion.value.ops,
        {
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          at: deletion.value.selection,
          newBlockId: intent.newBlockId,
          newHalf: SPLIT_HALVES.FIRST,
          ...allocationFields,
        },
      ];
      selection = { ...deletion.value.selection, offset: 0, zeroWidthBefore: 0 };
      break;
    }
    case "joinParagraphs": {
      const at = { story: intent.story, blockId: intent.blockId, offset: 0 };
      const paragraph = paragraphAt(document, at);
      const ownInsertedMark =
        mode.type === "suggesting" &&
        (paragraph?.pPrMark?.kind === "ins" || paragraph?.pPrMark?.kind === "moveTo") &&
        paragraph.pPrMark.info.author === mode.revision.author;
      const join = {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: intent.story,
        blockId: intent.blockId,
        nextBlockId: intent.nextBlockId,
        sectionBoundary: SECTION_BOUNDARY_POLICIES.REMOVE,
        ...(ownInsertedMark ? {} : tracked),
      } as const;
      if (mode.type === "suggesting") {
        const groups = editorParagraphGroups(document, intent.story);
        const firstGroup = groups.find(({ paragraphs }) =>
          paragraphs.some(({ paraId }) => idKey(paraId ?? "") === idKey(intent.blockId)),
        );
        const followingGroup = groups.find(({ paragraphs }) =>
          paragraphs.some(({ paraId }) => idKey(paraId ?? "") === idKey(intent.nextBlockId)),
        );
        const plan = createTrackedPlan({ document, revision: mode.revision, newIds: mode.newIds });
        const joined = plan.append(join);
        if (joined.isErr()) return Result.err(joined.error);
        if (ownInsertedMark) {
          // Cancelling our inserted boundary is physical, but its formatting
          // effect remains a suggestion on the original trailing paragraph.
          const before = paragraphAt(document, {
            story: intent.story,
            blockId: intent.nextBlockId,
            offset: 0,
          });
          const after = paragraphAt(plan.document(), {
            story: intent.story,
            blockId: intent.nextBlockId,
            offset: 0,
          });
          if (!before || !after) panic("An own-mark join lost its trailing paragraph.");
          const restored = plan.append({
            type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
            story: intent.story,
            blockId: intent.nextBlockId,
            patch: Object.fromEntries([
              ...Object.keys(after.formatting ?? {}).map((key) => [key, null]),
              ...Object.entries(before.formatting ?? {}),
            ]),
            whenEmpty:
              before.formatting === undefined ? EMPTY_PROPERTY_SETS.OMIT : EMPTY_PROPERTY_SETS.KEEP,
          });
          if (restored.isErr()) return Result.err(restored.error);
        }
        if (!firstGroup || !followingGroup)
          panic("A valid canonical join must belong to paragraph groups.");
        const source =
          firstGroup.text.length === 0
            ? followingGroup.paragraphs.at(-1)
            : firstGroup.paragraphs.at(-1);
        const desired = paragraphPropertiesOf(source?.formatting);
        const survivor = paragraphAt(plan.document(), {
          story: intent.story,
          blockId: followingGroup.blockId,
          offset: 0,
        });
        if (!survivor) panic("A valid canonical join must retain its following group survivor.");
        const current = paragraphPropertiesOf(survivor.formatting);
        if (!structurallyEqual(current, desired)) {
          const patched = plan.append({
            type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
            story: intent.story,
            blockId: followingGroup.blockId,
            patch: Object.fromEntries([
              ...Object.keys(current ?? {}).map((key) => [key, null]),
              ...Object.entries(desired ?? {}),
            ]),
            whenEmpty:
              source?.formatting === undefined
                ? EMPTY_PROPERTY_SETS.OMIT
                : EMPTY_PROPERTY_SETS.KEEP,
            revision: mode.revision,
            propertyReview: PROPERTY_REVIEW_POLICIES.APPEND,
          });
          if (patched.isErr()) return Result.err(patched.error);
        }
        ops = plan.ops;
      } else {
        ops = [join];
      }
      let offset = 0;
      if ((mode.type === "editing" || ownInsertedMark) && paragraph !== undefined) {
        offset = paragraphLength(paragraph);
      }
      selection = { story: intent.story, blockId: intent.nextBlockId, offset };
      break;
    }
    default: {
      const unreachable: never = intent;
      return unreachable;
    }
  }
  return Result.ok({ ops, selection });
};
