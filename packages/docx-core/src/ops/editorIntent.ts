/** One editor intent, compiled to direct or tracked document operations. */
import { Result, panic } from "better-result";
import { applyDocumentOps } from "./apply";
import { captureDocumentOp } from "./wire";
import { applyFormattingPatch } from "./patch";
import { paragraphNumberingReference } from "../model/paragraphNumbering";

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
import { compareGaps, defaultInsertionGap, leafSpans, zeroWidthLeavesAt } from "./leaves";
import { gapAfterInserted } from "./inline";
import { paragraphLength, paragraphLogicalText } from "./offsets";
import {
  appendTrackedDeletion,
  createTrackedPlan,
  selectedParagraphRuns,
  replacementDeletionSegments,
} from "./plan";
import { planTrackedReplace, rangeStartAfterDeletion } from "./rangeReplacement";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { structurallyEqual } from "./equality";
import { isRemovedRevisionNode, paragraphPropertiesOf, reviewFieldsOf } from "./review";
import {
  DOCUMENT_OP_TYPES,
  SECTION_BOUNDARY_POLICIES,
  PROPERTY_REVIEW_POLICIES,
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
  | { type: "splitParagraph"; at: TextPosition; to?: TextPosition; newBlockId: string }
  | { type: "joinParagraphs"; story: OpStory; blockId: string; nextBlockId: string };

/** Revision metadata and every fresh identity are supplied before compilation. */
export type EditorIntentMode =
  | { type: "editing"; newIds?: NewIds }
  | { type: "suggesting"; revision: RevisionStamp; newIds: NewIds };

export type CompileEditorIntentOptions = {
  intent: EditorIntent;
  mode: EditorIntentMode;
  /** Lowest candidate for compiler-owned pasted block IDs, including retired session IDs. */
  firstBlockId?: number;
};
export type CompiledEditorIntent = { ops: DocumentOp[]; selection: TextPosition };

/** Fresh identities for one input, bounded by its paragraph leaves and ancestor records. */
export const allocateEditorIntentIds = (document: Document, intent?: EditorIntent) => {
  const identities = packageIdentityKeys(document.package).concat(
    reservedIdentityKeysIn(document.package),
  );
  const paragraphs = (() => {
    if (intent === undefined) return storyParagraphs(document.package.document);
    switch (intent.type) {
      case "replaceFragment":
      case "replaceText":
      case "insertAtom":
      case "formatRun": {
        const selected = selectedParagraphRuns(document, intent.from, intent.to);
        return selected.isOk() ? selected.value.flat() : [];
      }
      case "moveFragment": {
        const selected = selectedParagraphRuns(document, intent.from, intent.to);
        return [
          ...(selected.isOk() ? selected.value.flat() : []),
          ...storyParagraphs(storyBody(document, intent.target.story)).filter(
            ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(intent.target.blockId),
          ),
        ];
      }
      case "splitParagraph": {
        if (intent.to !== undefined) {
          const selected = selectedParagraphRuns(document, intent.at, intent.to);
          return selected.isOk() ? selected.value.flat() : [];
        }
        return storyParagraphs(storyBody(document, intent.at.story)).filter(
          ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(intent.at.blockId),
        );
      }
      case "formatParagraph":
        return storyParagraphs(storyBody(document, intent.at.story)).filter(
          ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(intent.at.blockId),
        );
      case "setList": {
        const ids = new Set(intent.items.map(({ at }) => idKey(at.blockId)));
        const first = intent.items.at(0);
        return first === undefined
          ? []
          : storyParagraphs(storyBody(document, first.at.story)).filter(({ paragraph }) =>
              ids.has(idKey(paragraph.paraId ?? "")),
            );
      }
      case "joinParagraphs":
        return storyParagraphs(storyBody(document, intent.story)).filter(
          ({ paragraph }) =>
            idKey(paragraph.paraId ?? "") === idKey(intent.blockId) ||
            idKey(paragraph.paraId ?? "") === idKey(intent.nextBlockId),
        );
      default: {
        const exhaustive: never = intent;
        return exhaustive;
      }
    }
  })();
  // Each leaf can start a deletion segment; each ancestor can be cut at both
  // endpoints. Paragraph joins need a mark and a property-change stamp.
  const incoming =
    intent?.type === "replaceFragment" || intent?.type === "moveFragment" ? intent.paragraphs : [];
  const incomingDemand =
    incoming.length === 0
      ? 0
      : 8 +
        incoming.reduce(
          (total, paragraph) =>
            total +
            4 +
            leafSpans(paragraph.content).reduce(
              (count, span) => count + 4 * (1 + span.ancestors.length),
              0,
            ),
          0,
        );
  const demand =
    1 +
    incomingDemand +
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

type IntentRunFormattingOptions = {
  from: TextPosition;
  to: TextPosition;
  runProps?: TextFormatting;
  runPropsPatch?: RunPropsPatch;
};

const intentRunFormatting = (
  document: Document,
  { from, to, runProps, runPropsPatch }: IntentRunFormattingOptions,
): TextFormatting => {
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
  const newIds = {
    revision: mode.newIds.revision?.filter(
      (id) => !occupied.has(`${IDENTITY_SPACES.REVISION}:${id}`),
    ),
    control: mode.newIds.control?.filter((id) => !occupied.has(`${IDENTITY_SPACES.CONTROL}:${id}`)),
  };
  if (mode.type === "editing") return { type: "editing", newIds };
  if (!occupied.has(`${IDENTITY_SPACES.REVISION}:${mode.revision.id}`))
    return { type: "suggesting", revision: mode.revision, newIds };
  const id = newIds.revision?.at(0) ?? MAX_REVISION_ID + 1;
  return {
    type: "suggesting",
    revision: { ...mode.revision, id },
    newIds: { ...newIds, revision: newIds.revision?.slice(1) },
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
          from: intent.from,
          to: intent.to,
          revision: mode.revision,
          newIds: mode.newIds,
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
      // Open edges and retained suffixes explicitly carry destination
      // properties; closed pasted marks carry their source properties. A
      // range deletion may have joined paragraphs with different formatting.
      if (tail !== undefined) {
        const applied = applyDocumentOps(document, ops);
        if (applied.isErr()) return Result.err(applied.error);
        const current = paragraphAt(applied.value.document, selection);
        const patch = Object.fromEntries([
          ...Object.keys(current?.formatting ?? {}).map((key) => [key, null]),
          ...Object.entries(tail.formatting ?? {}),
        ]);
        const formatting = compileEditorIntent(applied.value.document, {
          intent: { type: "formatParagraph", at: selection, patch },
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
        const changes = [...(current.propertyChanges ?? [])];
        for (const change of paragraph.propertyChanges ?? []) {
          if (!changes.some((existing) => existing.info.id === change.info.id))
            changes.push(change);
        }
        if (
          paragraph.pPrMark !== undefined ||
          changes.length > (current.propertyChanges?.length ?? 0)
        ) {
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
      const deleted = compileEditorIntent(document, {
        intent: { type: "replaceText", from: intent.from, to: intent.to, text: "" },
        mode,
      });
      if (deleted.isErr()) return deleted;
      const applied = applyDocumentOps(document, deleted.value.ops);
      if (applied.isErr()) return Result.err(applied.error);
      let target = intent.target;
      for (const op of deleted.value.ops) {
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
      const inserted = compileEditorIntent(applied.value.document, {
        intent: {
          type: "replaceFragment",
          from: target,
          to: target,
          paragraphs: intent.paragraphs,
          openStart: intent.openStart,
          openEnd: intent.openEnd,
        },
        mode: modeAfterOps(applied.value.document, mode),
        ...(firstBlockId === undefined ? {} : { firstBlockId }),
      });
      if (inserted.isErr()) return inserted;
      ops = [...deleted.value.ops, ...inserted.value.ops];
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
        numId <= 0 ||
        numId > MAX_REVISION_ID ||
        intent.items.some(({ ilvl }) => !Number.isInteger(ilvl) || ilvl < 0 || ilvl > 8)
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
            ...tracked,
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
          ...tracked,
        },
      ];
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
                  formatting: intentRunFormatting(document, intent),
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
          formatting: intentRunFormatting(document, intent),
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
            selection = { ...op.at, offset: op.at.offset + text.length, zeroWidthBefore: 0 };
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
      selection =
        text === "" ? at : { ...at, offset: from.offset + text.length, zeroWidthBefore: 0 };
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
  return Result.ok({ ops: ops.map(captureDocumentOp), selection });
};
