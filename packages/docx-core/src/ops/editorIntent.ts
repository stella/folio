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
import {
  IDENTITY_SPACES,
  idKey,
  packageIdentityKeys,
  reservedIdentityKeysIn,
  packageParagraphIds,
} from "./ids";
import {
  compareGaps,
  defaultInsertionGap,
  isCommentAnchor,
  leafSpans,
  zeroWidthLeavesAt,
} from "./leaves";
import { deleteBetween, gapAfterInserted } from "./inline";
import { createCensusReader } from "./editorIntentCensus";
import { alikeDepth } from "./leaves";
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
import { isRemovedRevisionNode, paragraphPropertiesOf, reviewFieldsOf } from "./review";
import { runsMergeable } from "./runMerge";
import {
  DOCUMENT_OP_TYPES,
  PROPERTY_REVIEW_POLICIES,
  SECTION_BOUNDARY_POLICIES,
  EMPTY_PROPERTY_SETS,
  SPLIT_HALVES,
  type DocumentOp,
  type DeleteRangeOp,
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
      type: "replaceFragment";
      from: TextPosition;
      to: TextPosition;
      paragraphs: readonly Paragraph[];
      openStart: 0 | 1;
      openEnd: 0 | 1;
    }
  | {
      type: "moveFragment";
      from: TextPosition;
      to: TextPosition;
      target: TextPosition;
      paragraphs: readonly Paragraph[];
      openStart: 0 | 1;
      openEnd: 0 | 1;
    }
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

type CompileEditorIntentOptions = {
  intent: EditorIntent;
  mode: EditorIntentMode;
  /** Lowest candidate for compiler-owned pasted block IDs, including retired session IDs. */
  firstBlockId?: number;
};
type CompiledEditorIntent = { ops: DocumentOp[]; selection: TextPosition };

/** A split's paragraph identity is allocated with its other fresh identities. */
type EditorIntentAllocation = EditorIntent | SplitParagraphIntent;

const intentEndpoints = (intent: EditorIntentAllocation) => {
  switch (intent.type) {
    case "replaceFragment":
    case "moveFragment":
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
    const incoming =
      intent.type === "replaceFragment" || intent.type === "moveFragment" ? intent.paragraphs : [];
    for (const paragraph of incoming) {
      demand += 4;
      for (const span of leafSpans(paragraph.content)) demand += 4 * (1 + span.ancestors.length);
    }
    if (incoming.length > 0) demand += 8;
    if (intent.type === "moveFragment") {
      const target = storyParagraphs(storyBody(document, intent.target.story)).find(
        ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(intent.target.blockId),
      );
      if (target !== undefined) {
        demand += 3;
        for (const span of leafSpans(target.paragraph.content))
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

/** A compound intent consumes one identity pool across its sequential plans. */
const modeAfterOps = (document: Document, mode: EditorIntentMode): EditorIntentMode => {
  if (mode.newIds === undefined) return mode;
  const occupied = new Set(
    packageIdentityKeys(document.package).concat(reservedIdentityKeysIn(document.package)),
  );
  const revisionIds = mode.newIds.revision?.filter(
    (id) => !occupied.has(`${IDENTITY_SPACES.REVISION}:${id}`),
  );
  const controlIds = mode.newIds.control?.filter(
    (id) => !occupied.has(`${IDENTITY_SPACES.CONTROL}:${id}`),
  );
  const newIds = {
    ...(revisionIds === undefined ? {} : { revision: revisionIds }),
    ...(controlIds === undefined ? {} : { control: controlIds }),
  };
  if (mode.type === "editing") return { type: "editing", newIds };
  if (!occupied.has(`${IDENTITY_SPACES.REVISION}:${mode.revision.id}`))
    return { type: "suggesting", revision: mode.revision, newIds };
  const id = newIds.revision?.at(0) ?? MAX_REVISION_ID + 1;
  return {
    type: "suggesting",
    revision: { ...mode.revision, id },
    newIds: { ...newIds, ...(revisionIds === undefined ? {} : { revision: revisionIds.slice(1) }) },
  };
};

/** Imported identities are package-local facts, never destination identities. */
const isCopiedRangeStart = (value: object): boolean =>
  "type" in value &&
  (value.type === "bookmarkStart" ||
    value.type === "moveFromRangeStart" ||
    value.type === "moveToRangeStart");

const isCopiedRangeMarker = (value: object): boolean =>
  isCopiedRangeStart(value) ||
  ("type" in value &&
    (value.type === "bookmarkEnd" ||
      value.type === "moveFromRangeEnd" ||
      value.type === "moveToRangeEnd"));

type IdentifyClipboardParagraphsOptions = {
  document: Document;
  paragraphs: readonly Paragraph[];
  mode: EditorIntentMode;
  firstBlockId: number | undefined;
};

const identifyClipboardParagraphs = ({
  document,
  paragraphs,
  mode,
  firstBlockId,
}: IdentifyClipboardParagraphsOptions): Result<Paragraph[], DocumentOpRefusal> => {
  const copy = structuredClone([...paragraphs]);
  const occupied = new Set(
    packageIdentityKeys(document.package).concat(reservedIdentityKeysIn(document.package)),
  );
  for (const id of mode.newIds?.revision ?? []) occupied.add(`${IDENTITY_SPACES.REVISION}:${id}`);
  for (const id of mode.newIds?.control ?? []) occupied.add(`${IDENTITY_SPACES.CONTROL}:${id}`);
  if (mode.type === "suggesting") occupied.add(`${IDENTITY_SPACES.REVISION}:${mode.revision.id}`);
  const blockIds = new Set(packageParagraphIds(document.package).map(idKey));
  const markers = new Set<number>();
  const names = new Set<string>();
  const renamedBookmarks = new Map<string, string>();
  const visit = (value: unknown, callback: (value: object) => void): void => {
    if (typeof value !== "object" || value === null) return;
    if (Array.isArray(value)) {
      for (const item of value) visit(item, callback);
      return;
    }
    callback(value);
    for (const field of Object.values(value)) visit(field, callback);
  };
  visit(document.package, (value) => {
    if (
      "type" in value &&
      "id" in value &&
      typeof value.id === "number" &&
      isCopiedRangeMarker(value)
    )
      markers.add(value.id);
    if (isCopiedRangeStart(value) && "name" in value && typeof value.name === "string")
      names.add(value.name);
  });
  visit(copy, (value) => {
    if (!isCopiedRangeStart(value) || !("name" in value) || typeof value.name !== "string") return;
    const original = value.name;
    const renamed = renamedBookmarks.get(original);
    if (renamed !== undefined) {
      value.name = renamed;
      return;
    }
    let name = original;
    let suffix = 1;
    while (names.has(name)) {
      name = `${original}_paste${suffix}`;
      suffix += 1;
    }
    names.add(name);
    renamedBookmarks.set(original, name);
    value.name = name;
  });
  const remapped = new Map<string, number>();
  const nextInSpace = new Map<string, number>();
  let nextBlock = firstBlockId ?? 1;
  const freshBlock = () => {
    while (
      nextBlock <= 0x7fffffff &&
      blockIds.has(nextBlock.toString(16).padStart(8, "0").toUpperCase())
    )
      nextBlock += 1;
    if (nextBlock > 0x7fffffff) return undefined;
    const id = nextBlock.toString(16).padStart(8, "0").toUpperCase();
    blockIds.add(id);
    nextBlock += 1;
    return id;
  };
  const freshId = (space: string, oldId: number) => {
    const key = `${space}:${oldId}`;
    const existing = remapped.get(key);
    if (existing !== undefined) return existing;
    let id = nextInSpace.get(space) ?? 1;
    while (space === "bookmark" ? markers.has(id) : occupied.has(`${space}:${id}`)) id += 1;
    if (space === "bookmark") markers.add(id);
    else occupied.add(`${space}:${id}`);
    remapped.set(key, id);
    nextInSpace.set(space, id + 1);
    return id;
  };
  const remintedControls = new WeakSet<object>();
  visit(copy, (value) => {
    if (
      "type" in value &&
      value.type === "inlineSdt" &&
      "properties" in value &&
      typeof value.properties === "object" &&
      value.properties !== null &&
      "id" in value.properties &&
      typeof value.properties.id === "number"
    ) {
      value.properties.id = freshId(IDENTITY_SPACES.CONTROL, value.properties.id);
      remintedControls.add(value.properties);
    }
    if (
      "type" in value &&
      value.type === "hyperlink" &&
      "anchor" in value &&
      typeof value.anchor === "string"
    )
      value.anchor = renamedBookmarks.get(value.anchor) ?? value.anchor;
    if ("id" in value && typeof value.id === "number" && !remintedControls.has(value)) {
      if (
        "space" in value &&
        (value.space === IDENTITY_SPACES.REVISION || value.space === IDENTITY_SPACES.CONTROL)
      ) {
        // Cut provenance names the same copied records as info/properties IDs.
        value.id = freshId(value.space, value.id);
      } else if ("sdtType" in value) value.id = freshId(IDENTITY_SPACES.CONTROL, value.id);
      else if ("author" in value && !("type" in value))
        value.id = freshId(IDENTITY_SPACES.REVISION, value.id);
      else if (isCopiedRangeMarker(value)) value.id = freshId("bookmark", value.id);
    }
  });
  // Only split-off paragraphs need new IDs; the inline tail keeps its
  // destination identity, including when the session has exhausted block IDs.
  for (const [index, paragraph] of copy.entries()) {
    delete paragraph.paraId;
    if (index === copy.length - 1) continue;
    const id = freshBlock();
    if (id === undefined) {
      return Result.err(
        new DocumentOpRefusal({
          reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
          message: "The pasted block identity allocator is exhausted.",
          opType: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
        }),
      );
    }
    paragraph.paraId = id;
  }
  return Result.ok(copy);
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
  { intent, mode, firstBlockId }: CompileEditorIntentOptions,
): Result<CompiledEditorIntent, DocumentOpRefusal> => {
  if (
    firstBlockId !== undefined &&
    (!Number.isInteger(firstBlockId) || firstBlockId < 1 || firstBlockId > 0x80000000)
  ) {
    return Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
        message: "The pasted block identity allocator is exhausted or invalid.",
        opType: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      }),
    );
  }
  const allocationFields = mode.newIds === undefined ? {} : { newIds: mode.newIds };
  const tracked =
    mode.type === "suggesting"
      ? { revision: mode.revision, newIds: mode.newIds }
      : allocationFields;
  let ops: DocumentOp[];
  let selection: TextPosition;
  const editedSeams: TextPosition[] = [];
  switch (intent.type) {
    case "replaceFragment": {
      if (intent.paragraphs.length === 0)
        return compileEditorIntent(document, {
          intent: { type: "replaceText", from: intent.from, to: intent.to, text: "" },
          mode,
        });
      const destination = paragraphAt(document, intent.from);
      if (destination === undefined)
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
            message: "The clipboard destination paragraph does not exist.",
            opType: DOCUMENT_OP_TYPES.INSERT_CONTENT,
          }),
        );
      const endpoint = paragraphAt(document, intent.to);
      const destinationFormatting = splitParagraphFields(document, intent.from).formatting;
      const endpointFormatting = splitParagraphFields(document, intent.to).formatting;
      const prefixBoundary = intent.openStart === 0 && intent.from.offset > 0;
      const suffixBoundary =
        intent.openEnd === 0 &&
        endpoint !== undefined &&
        intent.to.offset < paragraphLength(endpoint);
      const normalized = [
        ...(prefixBoundary
          ? [
              {
                type: "paragraph",
                content: [],
                ...(destinationFormatting === undefined
                  ? {}
                  : { formatting: destinationFormatting }),
              } satisfies Paragraph,
            ]
          : []),
        ...intent.paragraphs,
        ...(suffixBoundary
          ? [
              {
                type: "paragraph",
                content: [],
                ...(endpointFormatting === undefined ? {} : { formatting: endpointFormatting }),
              } satisfies Paragraph,
            ]
          : []),
      ];
      const identified = identifyClipboardParagraphs({
        document,
        paragraphs: normalized,
        mode,
        firstBlockId,
      });
      if (identified.isErr()) return identified;
      const paragraphs = identified.value;
      for (const [index, paragraph] of paragraphs.entries()) {
        const openStart = index === 0 && intent.openStart === 1;
        const openEnd = index === normalized.length - 1 && intent.openEnd === 1;
        if (!openStart && !openEnd) continue;
        // The compiler owns these cloned paragraphs. Open edges join the
        // destination mark, keeping their authored inline content.
        delete paragraph.propertyChanges;
        delete paragraph.pPrMark;
        delete paragraph.reviewCarrier;
        delete paragraph.sectionProperties;
        delete paragraph.formatting;
        const inherited = openStart ? destinationFormatting : endpointFormatting;
        if (inherited !== undefined) paragraph.formatting = inherited;
      }
      if (paragraphs.some((paragraph) => paragraph.reviewCarrier !== undefined)) {
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            message: "A clipboard paragraph cannot import a private review-resolution carrier.",
            opType: DOCUMENT_OP_TYPES.INSERT_CONTENT,
          }),
        );
      }
      if (
        paragraphs.some(
          (paragraph) =>
            (paragraph.sectionProperties?.headerReferences?.length ?? 0) > 0 ||
            (paragraph.sectionProperties?.footerReferences?.length ?? 0) > 0,
        )
      ) {
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
            message:
              "Clipboard section headers and footers require importing their source package parts.",
            opType: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
          }),
        );
      }
      if (
        mode.type === "suggesting" &&
        paragraphs.some((paragraph) => paragraph.pPrMark !== undefined)
      ) {
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            message:
              "A copied paragraph mark cannot coexist with the tracked paste mark in the paragraph's single mark slot.",
            opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
          }),
        );
      }
      const tail = paragraphs.at(-1);
      if (
        mode.type === "suggesting" &&
        tail !== undefined &&
        (tail.sectionProperties !== undefined ||
          tail.pPrMark !== undefined ||
          (tail.propertyChanges?.length ?? 0) > 0)
      ) {
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            message:
              "A tracked paste cannot attach copied paragraph review or section metadata to a retained destination mark without a separately rejectable mark change.",
            opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
          }),
        );
      }
      const leading = paragraphs.slice(0, -1);
      if (mode.type === "suggesting") {
        const planned = planTrackedReplace(document, {
          sourceContainerPolicy: "separate",
          from: intent.from,
          to: intent.to,
          revision: mode.revision,
          newIds: mode.newIds,
          seamPolicy: INSERTION_SEAM_POLICIES.MERGE_PLAIN_RUNS,
          replacement: {
            paragraphs: leading.map(({ sectionProperties: _section, ...paragraph }) => paragraph),
            tail: { openStart: 0, openEnd: 0, content: tail?.content ?? [] },
          },
        });
        if (planned.isErr()) return Result.err(planned.error);
        ops = planned.value;
        const applied = applyDocumentOps(document, ops);
        if (applied.isErr()) return Result.err(applied.error);
        const at = rangeStartAfterDeletion({
          before: document,
          after: applied.value.document,
          from: intent.from,
          to: intent.to,
        });
        const insertionAt = leading.length > 0 ? { ...at, offset: 0, zeroWidthBefore: 0 } : at;
        const insertionParagraph = paragraphAt(document, insertionAt);
        selection = {
          ...insertionAt,
          ...gapAfterInserted(
            {
              offset: insertionAt.offset,
              zeroWidthBefore:
                insertionAt.zeroWidthBefore ??
                (insertionParagraph === undefined
                  ? 0
                  : defaultInsertionGap(insertionParagraph.content, insertionAt.offset)
                      .zeroWidthBefore),
            },
            tail?.content ?? [],
          ),
        };
      } else {
        const deleted = compileEditorIntent(document, {
          intent: { type: "replaceText", from: intent.from, to: intent.to, text: "" },
          mode,
        });
        if (deleted.isErr()) return deleted;
        ops = [...deleted.value.ops];
        let at = {
          ...deleted.value.selection,
          zeroWidthBefore:
            deleted.value.selection.zeroWidthBefore ??
            defaultInsertionGap(destination.content, intent.from.offset).zeroWidthBefore,
        };
        for (const paragraph of leading) {
          if (paragraph.content.length > 0)
            ops.push({
              type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
              at,
              slice: { openStart: 0, openEnd: 0, content: paragraph.content },
              seamPolicy: INSERTION_SEAM_POLICIES.MERGE_PLAIN_RUNS,
              ...allocationFields,
            });
          const {
            type: _type,
            paraId,
            content: _content,
            sectionProperties: _section,
            pPrMark: _mark,
            ...newParagraph
          } = paragraph;
          ops.push({
            type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
            at: {
              ...at,
              ...gapAfterInserted(
                { offset: at.offset, zeroWidthBefore: at.zeroWidthBefore ?? 0 },
                paragraph.content,
              ),
            },
            newBlockId: paraId ?? "",
            newHalf: SPLIT_HALVES.FIRST,
            newParagraph,
            ...allocationFields,
          });
          at = { story: at.story, blockId: at.blockId, offset: 0, zeroWidthBefore: 0 };
        }
        if (tail !== undefined && tail.content.length > 0)
          ops.push({
            type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
            at,
            slice: { openStart: 0, openEnd: 0, content: tail.content },
            ...allocationFields,
          });
        selection = {
          ...at,
          ...gapAfterInserted(
            { offset: at.offset, zeroWidthBefore: at.zeroWidthBefore ?? 0 },
            tail?.content ?? [],
          ),
        };
      }
      if (mode.type === "editing" && intent.openStart === 1 && leading.length > 0) {
        const first = leading.at(0);
        if (first?.paraId === undefined)
          panic("An identified leading clipboard paragraph must have an id.");
        editedSeams.push({
          story: intent.from.story,
          blockId: first.paraId,
          offset: intent.from.offset,
        });
      }
      // Open edges and retained suffixes explicitly carry destination
      // properties; closed pasted marks carry their source properties. A
      // range deletion may have joined paragraphs with different formatting.
      if (tail !== undefined) {
        const applied = applyDocumentOps(document, ops);
        if (applied.isErr()) return Result.err(applied.error);
        const trailing = editorParagraphGroups(applied.value.document, selection.story)
          .find(({ paragraphs: group }) =>
            group.some(({ paraId }) => idKey(paraId ?? "") === idKey(selection.blockId)),
          )
          ?.paragraphs.at(-1);
        const trailingMarkAt =
          trailing?.paraId === undefined
            ? selection
            : { ...selection, blockId: trailing.paraId, offset: 0 };
        const current = paragraphAt(applied.value.document, trailingMarkAt);
        const patch = Object.fromEntries([
          ...Object.keys(current?.formatting ?? {}).map((key) => [key, null]),
          ...Object.entries(tail.formatting ?? {}),
        ]);
        const formatting = compileEditorIntent(applied.value.document, {
          intent: { type: "formatParagraph", at: trailingMarkAt, patch },
          mode: modeAfterOps(applied.value.document, mode),
        });
        if (formatting.isErr()) return formatting;
        ops.push(...formatting.value.ops);
      }
      // Closed pasted marks are applied after structural allocation, when
      // their exact destination paragraph and existing review are known.
      for (const [index, paragraph] of paragraphs.entries()) {
        if (
          paragraph.pPrMark === undefined &&
          (paragraph.propertyChanges?.length ?? 0) === 0 &&
          paragraph.sectionProperties === undefined
        )
          continue;
        const blockId =
          index === paragraphs.length - 1 ? selection.blockId : (paragraph.paraId ?? "");
        const applied = applyDocumentOps(document, ops);
        if (applied.isErr()) return Result.err(applied.error);
        const current = paragraphAt(applied.value.document, {
          story: selection.story,
          blockId,
          offset: 0,
        });
        if (current === undefined) panic("An allocated clipboard paragraph must remain present.");
        if (paragraph.pPrMark !== undefined && current.pPrMark !== undefined) {
          return Result.err(
            new DocumentOpRefusal({
              reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
              message:
                "A copied paragraph mark cannot replace the destination's pending paragraph mark.",
              opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
            }),
          );
        }
        const expected = reviewFieldsOf(current);
        const copiedChange = paragraph.propertyChanges?.at(0);
        const currentChange = current.propertyChanges?.at(0);
        if (
          copiedChange !== undefined &&
          currentChange !== undefined &&
          packageIdentityKeys(document.package).includes(
            `${IDENTITY_SPACES.REVISION}:${currentChange.info.id}`,
          )
        )
          return Result.err(
            new DocumentOpRefusal({
              reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
              message: "Copied paragraph review cannot replace a pre-existing destination review.",
              opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
            }),
          );
        const changes = [...(paragraph.propertyChanges ?? current.propertyChanges ?? [])];
        if (paragraph.pPrMark !== undefined || copiedChange !== undefined) {
          const review = reviewFieldsOf(current);
          if (changes.length > 0) review.propertyChanges = changes;
          if (paragraph.pPrMark !== undefined) review.pPrMark = paragraph.pPrMark;
          ops.push({
            type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
            story: selection.story,
            blockId,
            expected,
            review,
          });
        }
        if (paragraph.sectionProperties !== undefined) {
          ops.push({
            type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
            endpoint: { type: "paragraph", blockId },
            expected:
              current.sectionProperties === undefined
                ? { type: Object.hasOwn(current, "sectionProperties") ? "undefined" : "omitted" }
                : { type: "present", value: current.sectionProperties },
            properties: { type: "present", value: paragraph.sectionProperties },
          });
        }
      }
      if (mode.type === "editing" && tail !== undefined) {
        const width = paragraphLength(tail);
        if (width > 0)
          for (const offset of new Set([selection.offset, selection.offset - width]))
            if (offset >= 0) editedSeams.push({ ...selection, offset });
      }
      // A closed paragraph payload ends outside that paragraph. Match the
      // forward text-selection affinity at its closing boundary when a next
      // untouched paragraph provides a text position; at document end the
      // caret remains at the payload's own last text position.
      if (
        intent.openEnd === 0 &&
        endpoint !== undefined &&
        intent.to.offset === paragraphLength(endpoint)
      ) {
        const bodyParagraphs = storyParagraphs(storyBody(document, intent.to.story));
        const endLocation = bodyParagraphs.find(
          ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(intent.to.blockId),
        );
        const following = bodyParagraphs.find(
          (location) =>
            endLocation !== undefined &&
            sameBlockList(location.list, endLocation.list) &&
            location.index === endLocation.index + 1,
        );
        if (following?.paragraph.paraId !== undefined) {
          selection = {
            story: intent.to.story,
            blockId: following.paragraph.paraId,
            ...defaultInsertionGap(following.paragraph.content, 0),
          };
        }
      }
      break;
    }
    case "moveFragment": {
      const range = selectedParagraphRuns(document, intent.from, intent.to);
      if (range.isErr()) return Result.err(range.error);
      const sourceParagraph = paragraphAt(document, intent.from);
      const fromGap =
        intent.from.zeroWidthBefore ??
        (sourceParagraph === undefined
          ? 0
          : zeroWidthLeavesAt(sourceParagraph.content, intent.from.offset).length);
      const toGap = intent.to.zeroWidthBefore ?? 0;
      if (
        idKey(intent.from.blockId) === idKey(intent.to.blockId) &&
        intent.from.offset === intent.to.offset &&
        fromGap >= toGap
      ) {
        if (
          fromGap > toGap &&
          (intent.from.zeroWidthBefore !== undefined || intent.to.zeroWidthBefore !== undefined)
        ) {
          return Result.err(
            new DocumentOpRefusal({
              reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
              message: "The move source marker gaps are reversed.",
              opType: DOCUMENT_OP_TYPES.DELETE_RANGE,
            }),
          );
        }
        return Result.ok({ ops: [], selection: intent.target });
      }
      const locations = storyParagraphs(storyBody(document, intent.from.story));
      const indexOf = (at: TextPosition) =>
        locations.findIndex(({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(at.blockId));
      const first = indexOf(intent.from);
      const last = indexOf(intent.to);
      const targetIndex = indexOf(intent.target);
      const targetGap =
        intent.target.zeroWidthBefore ??
        defaultInsertionGap(
          paragraphAt(document, intent.target)?.content ?? [],
          intent.target.offset,
        ).zeroWidthBefore;
      const targetPosition = { offset: intent.target.offset, zeroWidthBefore: targetGap };
      if (
        intent.target.story === intent.from.story &&
        targetIndex >= first &&
        targetIndex <= last &&
        (targetIndex !== first ||
          compareGaps(targetPosition, { offset: intent.from.offset, zeroWidthBefore: fromGap }) >=
            0) &&
        (targetIndex !== last ||
          compareGaps(targetPosition, { offset: intent.to.offset, zeroWidthBefore: toGap }) <= 0)
      ) {
        // Moving into the selected range is an explicit no-op, including its edges.
        return Result.ok({ ops: [], selection: intent.target });
      }
      const anchorDeletions: DeleteRangeOp[] = [];
      for (const { paragraph } of range.value.flat()) {
        const blockId = paragraph.paraId ?? "";
        const from =
          idKey(blockId) === idKey(intent.from.blockId)
            ? { offset: intent.from.offset, zeroWidthBefore: fromGap }
            : { offset: 0, zeroWidthBefore: 0 };
        const to =
          idKey(blockId) === idKey(intent.to.blockId)
            ? { offset: intent.to.offset, zeroWidthBefore: toGap }
            : {
                offset: paragraphLength(paragraph),
                zeroWidthBefore: zeroWidthLeavesAt(paragraph.content, paragraphLength(paragraph))
                  .length,
              };
        for (const span of leafSpans(paragraph.content)) {
          if (
            !isCommentAnchor(span.node) ||
            span.ancestors.some(isRemovedRevisionNode) ||
            compareGaps(span.before, from) < 0 ||
            compareGaps(span.after, to) > 0
          )
            continue;
          anchorDeletions.push({
            type: DOCUMENT_OP_TYPES.DELETE_RANGE,
            from: { story: intent.from.story, blockId, ...span.before },
            to: { story: intent.from.story, blockId, ...span.after },
          });
        }
      }
      if (anchorDeletions.length > 0 && mode.type === "suggesting")
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            message: "A tracked move cannot transfer comment anchors.",
            opType: DOCUMENT_OP_TYPES.DELETE_RANGE,
          }),
        );
      const deleted = compileEditorIntent(document, {
        intent: { type: "replaceText", from: intent.from, to: intent.to, text: "" },
        mode,
      });
      if (deleted.isErr()) return deleted;
      const logicalDeletions =
        anchorDeletions.length === 0
          ? deleted.value.ops
          : [
              ...deleted.value.ops
                .filter((op) => op.type === DOCUMENT_OP_TYPES.DELETE_RANGE)
                .concat(anchorDeletions)
                .sort(
                  (left, right) =>
                    indexOf(right.from) - indexOf(left.from) ||
                    compareGaps(
                      {
                        offset: right.from.offset,
                        zeroWidthBefore: right.from.zeroWidthBefore ?? 0,
                      },
                      { offset: left.from.offset, zeroWidthBefore: left.from.zeroWidthBefore ?? 0 },
                    ),
                ),
              ...deleted.value.ops.filter((op) => op.type !== DOCUMENT_OP_TYPES.DELETE_RANGE),
            ];
      const anchors = new Set(anchorDeletions);
      const deletionOps: DocumentOp[] = [];
      let afterDeletion = document;
      for (const planned of logicalDeletions) {
        let op = planned;
        if (planned.type === DOCUMENT_OP_TYPES.DELETE_RANGE && anchors.has(planned)) {
          const paragraph = paragraphAt(afterDeletion, planned.from);
          if (paragraph === undefined) panic("A move anchor lost its selected paragraph.");
          op = {
            type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
            story: planned.from.story,
            blockId: planned.from.blockId,
            expected: paragraph.content,
            content: deleteBetween(
              paragraph.content,
              { offset: planned.from.offset, zeroWidthBefore: planned.from.zeroWidthBefore ?? 0 },
              { offset: planned.to.offset, zeroWidthBefore: planned.to.zeroWidthBefore ?? 0 },
            ).content,
          };
        }
        const applied = applyDocumentOps(afterDeletion, [op]);
        if (applied.isErr()) return Result.err(applied.error);
        afterDeletion = applied.value.document;
        deletionOps.push(op);
      }
      let target = intent.target;
      for (const op of logicalDeletions) {
        if (
          op.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
          op.revision === undefined &&
          idKey(target.blockId) === idKey(op.from.blockId) &&
          target.story === op.from.story &&
          target.offset >= op.to.offset
        ) {
          const source = paragraphAt(document, op.from);
          const deletionFromGap =
            op.from.zeroWidthBefore ??
            zeroWidthLeavesAt(source?.content ?? [], op.from.offset).length;
          const deletionToGap = op.to.zeroWidthBefore ?? 0;
          const deletionTargetGap =
            target.zeroWidthBefore ??
            defaultInsertionGap(source?.content ?? [], target.offset).zeroWidthBefore;
          const rebased = Object.assign({}, target, {
            offset: target.offset - (op.to.offset - op.from.offset),
          });
          if (target.offset === op.to.offset)
            rebased.zeroWidthBefore = deletionFromGap + deletionTargetGap - deletionToGap;
          target = rebased;
        }
        if (
          op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS &&
          op.revision === undefined &&
          idKey(target.blockId) === idKey(op.blockId) &&
          target.story === op.story
        )
          target = Object.assign({}, target, { blockId: op.nextBlockId });
        else if (
          op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS &&
          op.revision === undefined &&
          idKey(target.blockId) === idKey(op.nextBlockId) &&
          target.story === op.story
        ) {
          const prefix = paragraphAt(document, { story: op.story, blockId: op.blockId, offset: 0 });
          const remaining =
            idKey(op.blockId) === idKey(intent.from.blockId) ? intent.from.offset : 0;
          if (prefix !== undefined) {
            const rebased = Object.assign({}, target, { offset: target.offset + remaining });
            if (target.offset === 0 && idKey(op.blockId) === idKey(intent.from.blockId))
              rebased.zeroWidthBefore = fromGap + (target.zeroWidthBefore ?? 0);
            target = rebased;
          }
        }
      }
      const inserted = compileEditorIntent(afterDeletion, {
        intent: {
          type: "replaceFragment",
          from: target,
          to: target,
          paragraphs: intent.paragraphs,
          openStart: intent.openStart,
          openEnd: intent.openEnd,
        },
        mode: modeAfterOps(afterDeletion, mode),
        ...(firstBlockId === undefined ? {} : { firstBlockId }),
      });
      if (inserted.isErr()) return inserted;
      ops = [...deletionOps, ...inserted.value.ops];
      selection = inserted.value.selection;
      break;
    }
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
          sourceContainerPolicy: "join",
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
          sourceContainerPolicy: "join",
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
      if (
        from.blockId !== to.blockId ||
        from.offset !== to.offset ||
        (from.zeroWidthBefore ?? 0) !== (to.zeroWidthBefore ?? 0)
      ) {
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
    const insertedWidth = intent.type === "replaceText" ? intent.text.length : 0;
    if (insertedWidth > 0) {
      for (const offset of new Set([selection.offset, selection.offset - insertedWidth]))
        if (offset >= 0) editedSeams.push({ ...selection, offset });
    }
    // Merge only the edited seams, keeping every authored interior run boundary.
    for (const at of editedSeams) {
      const paragraph = paragraphAt(current, at);
      if (paragraph === undefined) panic("An edited paragraph must exist at its seam.");
      const depth = textSeamDepth(paragraph, at.offset);
      if (depth === 0) continue;
      const join = { type: DOCUMENT_OP_TYPES.JOIN_INLINE, at, depth } as const;
      const joined = applyDocumentOp(current, join);
      if (joined.isErr()) return Result.err(joined.error);
      compact.push(join);
      current = joined.value.document;
    }
    ops = compact;
  }
  return Result.ok({ ops: ops.map(captureDocumentOp), selection });
};
