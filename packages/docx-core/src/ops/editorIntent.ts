/** One editor intent, compiled to direct or tracked document operations. */
import { INSERTION_SEAM_POLICIES } from "../model/content";
import { Result, panic } from "better-result";
import { applyDocumentOp, applyDocumentOps } from "./apply";
import { captureDocumentOp } from "./wire";
import { applyFormattingPatch } from "./patch";
import { isNumberingLevel } from "../model/numberingLevel";
import {
  isNumberingReference,
  NO_NUMBERING_NUM_ID,
  paragraphNumberingReference,
} from "../model/paragraphNumbering";

import {
  MAX_REVISION_ID,
  type Document,
  type Paragraph,
  type ParagraphContent,
  type TextFormatting,
  type TabContent,
  type BreakContent,
  type NumberingInstance,
  type AbstractNumbering,
} from "../model/document";
import { sameBlockList, storyBody, storyParagraphs } from "./blocks";
import { IDENTITY_SPACES, idKey } from "./ids";
import { createCensusReader } from "./editorIntentCensus";
import { defaultInsertionGap, alikeDepth, leafSpans, zeroWidthLeavesAt } from "./leaves";
import { paragraphLength, paragraphLogicalText } from "./offsets";
import {
  appendTrackedDeletion,
  createTrackedPlan,
  selectedParagraphRuns,
  replacementDeletionSegments,
  filterNewIds,
  trimAppliedNewIds,
} from "./plan";
import { planTrackedReplace, rangeStartAfterDeletion } from "./rangeReplacement";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { structurallyEqual } from "./equality";
import { runsMergeable } from "./runMerge";
import { isRemovedRevisionNode, paragraphPropertiesOf } from "./review";
import {
  DOCUMENT_OP_TYPES,
  PROPERTY_REVIEW_POLICIES,
  SECTION_BOUNDARY_POLICIES,
  EMPTY_PROPERTY_SETS,
  SPLIT_HALVES,
  type DocumentOp,
  type NewIds,
  type OpStory,
  type RevisionStamp,
  type TextPosition,
  type RunPropsPatch,
  type ParagraphPropsPatch,
} from "./types";

type SplitParagraphIntent = {
  type: "splitParagraph";
  at: TextPosition;
  to?: TextPosition;
};

/** Positions use canonical physical offsets, including retained deleted content. */
export type EditorIntent =
  | {
      type: "replaceText";
      from: TextPosition;
      to: TextPosition;
      text: string;
      runProps?: TextFormatting;
      runPropsPatch?: RunPropsPatch;
    }
  | { type: "formatRun"; from: TextPosition; to: TextPosition; patch: RunPropsPatch }
  | { type: "formatParagraph"; at: TextPosition; patch: ParagraphPropsPatch }
  | {
      type: "setList";
      items: readonly { at: TextPosition; ilvl: number }[];
      target:
        | { type: "existing"; numId: number }
        | { type: "new"; num: NumberingInstance; abstractNum?: AbstractNumbering };
    }
  | {
      type: "insertAtom";
      from: TextPosition;
      to: TextPosition;
      atom: TabContent | BreakContent;
      runProps?: TextFormatting;
      runPropsPatch?: RunPropsPatch;
    }
  | (SplitParagraphIntent & { newBlockId: string })
  | { type: "joinParagraphs"; story: OpStory; blockId: string; nextBlockId: string };

/** Revision metadata and every fresh identity are supplied before compilation. */
export type EditorIntentMode =
  | { type: "editing"; newIds?: NewIds }
  | { type: "suggesting"; revision: RevisionStamp; newIds: NewIds };

type CompileEditorIntentOptions = { intent: EditorIntent; mode: EditorIntentMode };
type CompiledEditorIntent = { ops: DocumentOp[]; selection: TextPosition };

/** A split's paragraph identity is allocated with its other fresh identities. */
type EditorIntentAllocation = EditorIntent | SplitParagraphIntent;

const intentEndpoints = (intent: EditorIntentAllocation) => {
  switch (intent.type) {
    case "replaceText":
    case "insertAtom":
    case "formatRun":
      return {
        story: intent.from.story,
        fromId: intent.from.blockId,
        toId: intent.to.blockId,
        fromOffset: intent.from.offset,
        toOffset: intent.to.offset,
      };
    case "formatParagraph":
      return {
        story: intent.at.story,
        fromId: intent.at.blockId,
        toId: intent.at.blockId,
        fromOffset: undefined,
        toOffset: undefined,
      };
    case "setList": {
      const first = intent.items.at(0);
      const last = intent.items.at(-1);
      if (first === undefined || last === undefined) panic("List allocation requires a paragraph.");
      return {
        story: first.at.story,
        fromId: first.at.blockId,
        toId: last.at.blockId,
        fromOffset: undefined,
        toOffset: undefined,
      };
    }
    case "splitParagraph":
      return {
        story: intent.at.story,
        fromId: intent.at.blockId,
        toId: (intent.to ?? intent.at).blockId,
        fromOffset: intent.at.offset,
        toOffset: (intent.to ?? intent.at).offset,
      };
    case "joinParagraphs":
      return {
        story: intent.story,
        fromId: intent.blockId,
        toId: intent.nextBlockId,
        fromOffset: undefined,
        toOffset: undefined,
      };
    default: {
      const unreachable: never = intent;
      return unreachable;
    }
  }
};

/** Cache only immutable document versions; never retain a retired version strongly. */
export const createEditorIntentIdAllocator = () => {
  const readCensus = createCensusReader();
  return (document: Document, intent: EditorIntentAllocation) => {
    const census = readCensus(document);
    const { story, fromId, toId, fromOffset, toOffset } = intentEndpoints(intent);
    let locations = census.stories.get(story);
    if (locations === undefined) {
      locations = storyParagraphs(storyBody(document, story));
      census.stories.set(story, locations);
    }
    const first = locations.findIndex(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(fromId),
    );
    const last = locations.findIndex(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(toId),
    );
    // Every selected leaf may start a deletion segment; ancestors may be cut
    // at both endpoints. Each paragraph adds join/mark/property stamps.
    let demand = 1;
    const selected = first < 0 || last < first ? [] : locations.slice(first, last + 1);
    for (const { paragraph } of selected) {
      demand += 3;
      const blockId = idKey(paragraph.paraId ?? "");
      for (const span of leafSpans(paragraph.content)) {
        if (blockId === idKey(fromId) && fromOffset !== undefined && span.after.offset < fromOffset)
          continue;
        if (blockId === idKey(toId) && toOffset !== undefined && span.before.offset > toOffset)
          continue;
        demand += 4 * (1 + span.ancestors.length);
      }
    }
    const fresh = (occupied: ReadonlySet<number>, count: number) => {
      const ids: number[] = [];
      for (let id = 1; id <= MAX_REVISION_ID && ids.length < count; id += 1) {
        if (!occupied.has(id)) ids.push(id);
      }
      return ids;
    };
    const revision = fresh(census.revisions, demand + 1);
    let nextParagraph = 1;
    while (census.paragraphs.has(nextParagraph.toString(16).padStart(8, "0").toUpperCase()))
      nextParagraph += 1;
    return {
      revisionId: revision.at(0) ?? MAX_REVISION_ID + 1,
      newBlockId: nextParagraph.toString(16).padStart(8, "0").toUpperCase(),
      newIds: { revision: revision.slice(1), control: fresh(census.controls, demand) },
    };
  };
};

/** Fresh identities bounded by the edit, with one package census per document version. */
export const allocateEditorIntentIds = createEditorIntentIdAllocator();

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
const formattingAt = (paragraph: Paragraph, offset: number): TextFormatting | undefined => {
  const spans = leafSpans(paragraph.content).filter(
    ({ ancestors }) => !ancestors.some(isRemovedRevisionNode),
  );
  const span =
    spans.find(({ before, after }) => before.offset <= offset && offset < after.offset) ??
    spans.findLast(({ after }) => after.offset <= offset) ??
    spans.find(({ before }) => before.offset > offset);
  return span?.ancestors.findLast((ancestor) => ancestor.type === "run")?.formatting;
};

const authoredFormatting = (
  document: Document,
  from: TextPosition,
  to: TextPosition,
): TextFormatting | undefined => {
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

type IntentRunFormattingOptions = {
  from: TextPosition;
  to: TextPosition;
  runProps?: TextFormatting;
  runPropsPatch?: RunPropsPatch;
};

const intentRunFormatting = (
  document: Document,
  { from, to, runProps, runPropsPatch }: IntentRunFormattingOptions,
) => {
  const authored = runProps ?? authoredFormatting(document, from, to);
  return runPropsPatch === undefined
    ? authored
    : (applyFormattingPatch(authored, runPropsPatch) ?? {});
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

type JoinParagraphPropertyAlignmentOptions = {
  document: Document;
  current: Document;
  intent: Extract<EditorIntent, { type: "joinParagraphs" }>;
  revision?: RevisionStamp;
};

/** Direct and tracked joins choose paragraph properties from the same editing-view group. */
const joinParagraphPropertyAlignment = ({
  document,
  current,
  intent,
  revision,
}: JoinParagraphPropertyAlignmentOptions) => {
  const groups = editorParagraphGroups(document, intent.story);
  const firstGroup = groups.find(({ paragraphs }) =>
    paragraphs.some(({ paraId }) => idKey(paraId ?? "") === idKey(intent.blockId)),
  );
  const followingGroup = groups.find(({ paragraphs }) =>
    paragraphs.some(({ paraId }) => idKey(paraId ?? "") === idKey(intent.nextBlockId)),
  );
  if (!firstGroup || !followingGroup)
    panic("A valid canonical join must belong to paragraph groups.");
  const source =
    firstGroup.text.length === 0 ? followingGroup.paragraphs.at(-1) : firstGroup.paragraphs.at(-1);
  const survivor = paragraphAt(current, {
    story: intent.story,
    blockId: followingGroup.blockId,
    offset: 0,
  });
  if (!survivor) panic("A valid canonical join must retain its following group survivor.");
  const desired = paragraphPropertiesOf(source?.formatting);
  const existing = paragraphPropertiesOf(survivor.formatting);
  if (structurallyEqual(existing, desired)) return undefined;
  return {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: intent.story,
    blockId: followingGroup.blockId,
    patch: Object.fromEntries([
      ...Object.keys(existing ?? {}).map((key) => [key, null]),
      ...Object.entries(desired ?? {}),
    ]),
    whenEmpty:
      source?.formatting === undefined ? EMPTY_PROPERTY_SETS.OMIT : EMPTY_PROPERTY_SETS.KEEP,
    ...(revision === undefined ? {} : { revision }),
  } as const;
};

/** Join only plain runs meeting at an edited seam, within the same container. */
const textSeamDepth = (paragraph: Paragraph, offset: number): number => {
  const spans = leafSpans(paragraph.content);
  for (const [index, right] of spans.entries()) {
    const left = spans.at(index - 1);
    if (
      index === 0 ||
      left === undefined ||
      left.after.offset !== offset ||
      right.before.offset !== offset
    )
      continue;
    const leftRun = left.ancestors.at(-1);
    const rightRun = right.ancestors.at(-1);
    if (
      leftRun?.type !== "run" ||
      rightRun?.type !== "run" ||
      leftRun === rightRun ||
      leftRun.content.at(-1) !== left.node ||
      rightRun.content.at(0) !== right.node ||
      left.ancestors.length !== right.ancestors.length ||
      !left.ancestors
        .slice(0, -1)
        .every((parent, parentIndex) => parent === right.ancestors[parentIndex]) ||
      !runsMergeable(leftRun, rightRun)
    )
      continue;
    return alikeDepth(leftRun, rightRun);
  }
  return 0;
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
    case "setList": {
      const first = intent.items.at(0);
      if (first === undefined)
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
            message: "A list intent must address a paragraph.",
            opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          }),
        );
      const numId =
        intent.target.type === "existing" ? intent.target.numId : intent.target.num.numId;
      if (
        !Number.isInteger(numId) ||
        !isNumberingReference(numId) ||
        numId < NO_NUMBERING_NUM_ID ||
        numId > MAX_REVISION_ID ||
        intent.items.some(({ ilvl }) => !isNumberingLevel(ilvl))
      )
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
            message: "List references require a valid numbering id and level.",
            opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          }),
        );
      const creation =
        intent.target.type === "new"
          ? [
              {
                type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
                num: intent.target.num,
                ...(intent.target.abstractNum === undefined
                  ? {}
                  : { abstractNum: intent.target.abstractNum }),
              } as const,
            ]
          : [];
      const paragraphOps = intent.items.map(
        ({ at, ilvl }) =>
          ({
            type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
            story: at.story,
            blockId: at.blockId,
            patch: { numPr: paragraphNumberingReference({ numId, ilvl }) },
            ...(mode.type === "suggesting"
              ? { propertyReview: PROPERTY_REVIEW_POLICIES.APPEND }
              : {}),
            ...(mode.type === "suggesting" ? { revision: mode.revision } : {}),
          }) as const,
      );
      if (mode.type === "editing") ops = [...creation, ...paragraphOps];
      else {
        // Numbering definitions are package resources, not OOXML revision records.
        // Both modes allocate the same resource; the paragraph property is tracked.
        const created = applyDocumentOps(document, creation);
        if (created.isErr()) return Result.err(created.error);
        const plan = createTrackedPlan({
          document: created.value.document,
          revision: mode.revision,
          newIds: mode.newIds,
        });
        for (const op of paragraphOps) {
          const appended = plan.append(op);
          if (appended.isErr()) return Result.err(appended.error);
        }
        ops = [...creation, ...plan.ops];
      }
      selection = first.at;
      break;
    }
    case "formatParagraph": {
      ops = [
        {
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          story: intent.at.story,
          blockId: intent.at.blockId,
          patch: intent.patch,
          ...(mode.type === "suggesting"
            ? { propertyReview: PROPERTY_REVIEW_POLICIES.APPEND }
            : {}),
          ...(mode.type === "suggesting" ? { revision: mode.revision } : {}),
        },
      ];
      if (mode.type === "suggesting") {
        const plan = createTrackedPlan({ document, revision: mode.revision, newIds: mode.newIds });
        for (const op of ops) {
          if (op.type !== DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS)
            panic("A paragraph-format intent compiled to a different operation.");
          const appended = plan.append(op);
          if (appended.isErr()) return Result.err(appended.error);
        }
        ops = plan.ops;
      }
      selection = intent.at;
      break;
    }
    case "formatRun": {
      const partitioned = selectedParagraphRuns(document, intent.from, intent.to);
      if (partitioned.isErr()) return Result.err(partitioned.error);
      ops = partitioned.value.flatMap((run) =>
        run.map(({ paragraph }) => {
          const blockId = paragraph.paraId ?? "";
          return {
            type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
            from:
              idKey(blockId) === idKey(intent.from.blockId)
                ? intent.from
                : { story: intent.from.story, blockId, offset: 0 },
            to:
              idKey(blockId) === idKey(intent.to.blockId)
                ? intent.to
                : { story: intent.to.story, blockId, offset: paragraphLength(paragraph) },
            patch: intent.patch,
            ...(mode.type === "suggesting"
              ? { propertyReview: PROPERTY_REVIEW_POLICIES.APPEND }
              : {}),
            ...tracked,
          } as const;
        }),
      );
      if (mode.type === "suggesting") {
        const plan = createTrackedPlan({ document, revision: mode.revision, newIds: mode.newIds });
        for (const op of ops) {
          if (op.type !== DOCUMENT_OP_TYPES.SET_RUN_PROPS)
            panic("A run-format intent compiled to a different operation.");
          const appended = plan.append(op);
          if (appended.isErr()) return Result.err(appended.error);
        }
        ops = plan.ops;
      }
      selection = intent.to;
      break;
    }
    case "insertAtom": {
      const formatting = intentRunFormatting(document, intent);
      if (mode.type === "suggesting") {
        const planned = planTrackedReplace(document, {
          from: intent.from,
          to: intent.to,
          revision: mode.revision,
          newIds: mode.newIds,
          replacement: {
            paragraphs: [],
            tail: {
              openStart: 0,
              openEnd: 0,
              content: [
                {
                  type: "run",
                  ...(formatting === undefined ? {} : { formatting }),
                  content: [intent.atom],
                },
              ],
            },
          },
        });
        if (planned.isErr()) return Result.err(planned.error);
        ops = planned.value;
        const insertion = ops.findLast((op) => op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT);
        selection =
          insertion?.type === DOCUMENT_OP_TYPES.INSERT_CONTENT
            ? { ...insertion.at, offset: insertion.at.offset + 1 }
            : intent.from;
        break;
      }
      const deletion = compileEditorIntent(document, {
        intent: { type: "replaceText", from: intent.from, to: intent.to, text: "" },
        mode,
      });
      if (deletion.isErr()) return deletion;
      const at = deletion.value.selection;
      const content = [
        {
          type: "run",
          ...(formatting === undefined ? {} : { formatting }),
          content: [intent.atom],
        },
      ] satisfies ParagraphContent[];
      const insertion = {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        at,
        slice: {
          openStart: 0,
          openEnd: 0,
          content,
        },
        ...tracked,
      } as const satisfies DocumentOp;
      ops = [...deletion.value.ops, insertion];
      selection = { ...at, offset: at.offset + 1 };
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
      const runProps = intentRunFormatting(document, intent);
      // Both modes insert the same authored run; source XML attributes belong
      // to the existing run rather than to newly typed text.
      const content = (
        text === ""
          ? []
          : [
              {
                type: "run",
                ...(runProps === undefined ? {} : { formatting: runProps }),
                content: [{ type: "text", text }],
              },
            ]
      ) satisfies ParagraphContent[];
      if (mode.type === "suggesting") {
        const planned = planTrackedReplace(document, {
          from,
          to,
          revision: mode.revision,
          newIds: mode.newIds,
          seamPolicy: INSERTION_SEAM_POLICIES.MERGE_PLAIN_RUNS,
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
            selection = {
              ...op.at,
              offset: op.at.offset + text.length,
              ...(text.length > 0 && op.at.zeroWidthBefore !== undefined
                ? { zeroWidthBefore: 0 }
                : {}),
            };
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
            const start =
              idKey(blockId) === idKey(from.blockId)
                ? from
                : { story: from.story, blockId, offset: 0, zeroWidthBefore: 0 };
            const end =
              idKey(blockId) === idKey(to.blockId)
                ? to
                : {
                    story: from.story,
                    blockId,
                    offset: paragraphLength(item),
                    zeroWidthBefore: zeroWidthLeavesAt(item.content, paragraphLength(item)).length,
                  };
            const segments = replacementDeletionSegments({
              spans: leafSpans(item.content),
              from: {
                offset: start.offset,
                zeroWidthBefore:
                  start.zeroWidthBefore ?? zeroWidthLeavesAt(item.content, start.offset).length,
              },
              to: { offset: end.offset, zeroWidthBefore: end.zeroWidthBefore ?? 0 },
              mode: { type: "editing" },
            });
            for (const segment of segments.toReversed()) {
              ops.push({
                type: DOCUMENT_OP_TYPES.DELETE_RANGE,
                ...allocationFields,
                from: { story: from.story, blockId, ...segment.from },
                to: { story: from.story, blockId, ...segment.to },
              });
            }
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
      // Resolve insertion affinity in the input, before deletion shifts trailing
      // zero-width leaves onto this offset.
      const at = {
        ...from,
        blockId: survivorId,
        zeroWidthBefore:
          from.zeroWidthBefore ??
          defaultInsertionGap(paragraph.content, from.offset).zeroWidthBefore,
      };
      if (text !== "")
        ops.push({
          type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
          at,
          slice: { openStart: 0, openEnd: 0, content },
          ...allocationFields,
        });
      selection = {
        ...at,
        offset: from.offset + text.length,
        ...(text.length > 0 && at.zeroWidthBefore !== undefined ? { zeroWidthBefore: 0 } : {}),
      };
      break;
    }
    case "splitParagraph": {
      if (intent.to === undefined) {
        const split = {
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          at: intent.at,
          newBlockId: intent.newBlockId,
          newHalf: SPLIT_HALVES.FIRST,
          newParagraph: splitParagraphFields(document, intent.at),
          ...tracked,
        } as const;
        if (mode.type === "suggesting") {
          const plan = createTrackedPlan({
            document,
            revision: mode.revision,
            newIds: mode.newIds,
          });
          const appended = plan.append(split);
          if (appended.isErr()) return Result.err(appended.error);
          ops = plan.ops;
        } else ops = [split];
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
      const deleted = applyDocumentOps(document, deletion.value.ops);
      if (deleted.isErr()) return Result.err(deleted.error);
      ops = [
        ...deletion.value.ops,
        {
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          at: deletion.value.selection,
          newBlockId: intent.newBlockId,
          newHalf: SPLIT_HALVES.FIRST,
          newParagraph: splitParagraphFields(deleted.value.document, deletion.value.selection),
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
        const alignment = joinParagraphPropertyAlignment({
          document,
          current: plan.document(),
          intent,
          revision: mode.revision,
        });
        if (alignment !== undefined) {
          const patched = plan.append(alignment);
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
  // Simulate direct pieces in order: each consumes its own pool, and later
  // pieces must not reuse identities created by an earlier piece.
  if (mode.type === "editing") {
    let current = document;
    const compact: DocumentOp[] = [];
    const takenRevision = new Set<number>();
    const takenControl = new Set<number>();
    const { story } = intentEndpoints(intent);
    for (const input of ops) {
      const op = filterNewIds(
        input,
        (space, id) => !(space === IDENTITY_SPACES.REVISION ? takenRevision : takenControl).has(id),
      );
      const applied = applyDocumentOp(current, op);
      if (applied.isErr()) return Result.err(applied.error);
      const trimmed = trimAppliedNewIds({ op, applied: applied.value, story });
      if ("newIds" in trimmed) {
        for (const id of trimmed.newIds?.revision ?? []) takenRevision.add(id);
        for (const id of trimmed.newIds?.control ?? []) takenControl.add(id);
      }
      compact.push(trimmed);
      current = applied.value.document;
    }
    if (intent.type === "joinParagraphs") {
      const alignment = joinParagraphPropertyAlignment({ document, current, intent });
      if (alignment !== undefined) {
        const aligned = applyDocumentOp(current, alignment);
        if (aligned.isErr()) return Result.err(aligned.error);
        compact.push(alignment);
        current = aligned.value.document;
      }
    }
    if (intent.type === "replaceText" && intent.text.length > 0) {
      // Closed insertion slices preserve authored boundaries. Merge only the
      // two edited seams the parser would merge, with inverses in the journal.
      for (const offset of new Set([selection.offset, selection.offset - intent.text.length])) {
        const at = { ...selection, offset };
        const paragraph = paragraphAt(current, at);
        if (paragraph === undefined) panic("An edited paragraph must exist at its seam.");
        const depth = textSeamDepth(paragraph, offset);
        if (depth === 0) continue;
        const join = { type: DOCUMENT_OP_TYPES.JOIN_INLINE, at, depth } as const;
        const joined = applyDocumentOp(current, join);
        if (joined.isErr()) return Result.err(joined.error);
        compact.push(join);
        current = joined.value.document;
      }
    }
    ops = compact;
  }
  return Result.ok({ ops: ops.map(captureDocumentOp), selection });
};
