/**
 * Tracked changes: how an operation carrying a {@link RevisionStamp} records
 * its edit instead of making it, and the pieces resolution shares with it.
 *
 * A tracked insertion or deletion wraps the leaves it names in an `Insertion`
 * or `Deletion` carrying the stamp. The wrapper goes at the deepest level
 * that admits one: the paragraph, an inline content control, a transparent
 * wrapper, and for a deletion also an insertion or a move the range falls
 * inside. Records below that level which the range cuts (runs, hyperlinks,
 * and for an insertion another stamp's insertion) are cut at the range ends,
 * each piece keeping every field of the record, as a direct edit cuts them.
 */

import type {
  Paragraph,
  ParagraphContent,
  ParagraphFormatting,
  ParagraphPropertyChange,
  TrackedChangeInfo,
  TrackedRunContent,
} from "../model/document";
import { structurallyEqual } from "./equality";
import { IDENTITY_SPACES, slotKey } from "./ids";
import {
  childNodes,
  type Cursor,
  type Gap,
  isCommentAnchor,
  type InlineNode,
  isParagraphContent,
  mergeLists,
  partitionNode,
  rebuildNode,
  startCursor,
} from "./leaves";
import { isTrackedChild } from "./offsets";
import { DOCUMENT_OP_REFUSAL_REASONS, type DocumentOpRefusalReason } from "./refusal";
import {
  PARAGRAPH_MARK_FORMATTING_KEYS,
  type ParagraphReviewFields,
  type RevisionStamp,
} from "./types";

// ---------------------------------------------------------------------------
// Stamps
// ---------------------------------------------------------------------------

/** The tracked-change metadata a stamp writes on the records it creates. */
export const stampInfo = (stamp: RevisionStamp): TrackedChangeInfo => {
  const info: TrackedChangeInfo = { id: stamp.id, author: stamp.author, date: stamp.date };
  if (stamp.initials !== undefined) {
    info.initials = stamp.initials;
  }
  return info;
};

/** Whether a record's metadata is the stamp's, whatever its id. */
export const carriesStamp = (info: TrackedChangeInfo, stamp: RevisionStamp): boolean =>
  structurallyEqual({ ...info, id: 0 }, { ...stampInfo(stamp), id: 0 });

/** Tracked insertions, deletions and moves. */
export type TrackedWrapper = Extract<
  InlineNode,
  { type: "insertion" | "deletion" | "moveFrom" | "moveTo" }
>;

export const isTrackedWrapper = (node: InlineNode): node is TrackedWrapper =>
  node.type === "insertion" ||
  node.type === "deletion" ||
  node.type === "moveFrom" ||
  node.type === "moveTo";

/** Tracked content on its way in: an insertion or the destination of a move. */
export const isAddedRevision = (node: InlineNode): boolean =>
  node.type === "insertion" || node.type === "moveTo";

/** Tracked content on its way out: a deletion or the source of a move. */
export const isRemovedRevisionNode = (node: InlineNode): boolean =>
  node.type === "deletion" || node.type === "moveFrom";

/** The tracked-change metadata of every record a paragraph and its content carry. */
const revisionInfosOf = (paragraph: Paragraph): TrackedChangeInfo[] => {
  const out: TrackedChangeInfo[] = [];
  for (const change of paragraph.propertyChanges ?? []) out.push(change.info);
  if (paragraph.pPrMark !== undefined) out.push(paragraph.pPrMark.info);
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (isTrackedWrapper(node)) {
        out.push(node.info);
      } else if (node.type === "run") {
        for (const change of node.propertyChanges ?? []) out.push(change.info);
      }
      visit(childNodes(node) ?? []);
    }
  };
  visit(paragraph.content);
  return out;
};

/**
 * The revision ids of the records carrying the stamp that `known` does not
 * name (slot keys): the records an operation created, in document order.
 */
export const stampedRevisionIds = (
  paragraphs: readonly Paragraph[],
  stamp: RevisionStamp,
  known: ReadonlySet<string>,
): number[] => {
  const out: number[] = [];
  for (const paragraph of paragraphs) {
    for (const info of revisionInfosOf(paragraph)) {
      const key = slotKey({ space: IDENTITY_SPACES.REVISION, id: info.id });
      if (carriesStamp(info, stamp) && !known.has(key) && !out.includes(info.id)) {
        out.push(info.id);
      }
    }
  }
  return out;
};

// ---------------------------------------------------------------------------
// Paragraph properties a property change records
// ---------------------------------------------------------------------------

const MARK_KEYS: ReadonlySet<string> = new Set(PARAGRAPH_MARK_FORMATTING_KEYS);

/** Whether a patch or property set names a key of the paragraph mark's run properties. */
export const namesMarkFormatting = (formatting: object): boolean =>
  Object.entries(formatting).some(([key, value]) => value !== undefined && MARK_KEYS.has(key));

const pick = (
  formatting: ParagraphFormatting | undefined,
  keep: (key: string) => boolean,
): ParagraphFormatting | undefined => {
  if (formatting === undefined) {
    return undefined;
  }
  // SAFETY: the entries are a subset of a `ParagraphFormatting`'s own.
  return Object.fromEntries(
    Object.entries(formatting).filter(([key]) => keep(key)),
  ) as ParagraphFormatting;
};

/** The paragraph properties a `w:pPrChange` records: all but the mark's run properties. */
export const paragraphPropertiesOf = (
  formatting: ParagraphFormatting | undefined,
): ParagraphFormatting | undefined => pick(formatting, (key) => !MARK_KEYS.has(key));

/** The paragraph mark's run properties of a property set; `undefined` when it states none. */
const markFormattingOf = (
  formatting: ParagraphFormatting | undefined,
): ParagraphFormatting | undefined => {
  const marks = pick(formatting, (key) => MARK_KEYS.has(key));
  return marks === undefined || Object.keys(marks).length === 0 ? undefined : marks;
};

/** Whether two property sets state the same paragraph properties. */
export const sameParagraphProperties = (
  left: ParagraphFormatting | undefined,
  right: ParagraphFormatting | undefined,
): boolean =>
  structurallyEqual(paragraphPropertiesOf(left) ?? {}, paragraphPropertiesOf(right) ?? {});

/**
 * The paragraph properties of `properties` with the mark's run properties of
 * `marks`. `undefined` when the first is absent and the second states no mark
 * properties.
 */
export const withMarkFormatting = (
  properties: ParagraphFormatting | undefined,
  marks: ParagraphFormatting | undefined,
): ParagraphFormatting | undefined => {
  const markFormatting = markFormattingOf(marks);
  if (properties === undefined && markFormatting === undefined) {
    return undefined;
  }
  return { ...paragraphPropertiesOf(properties), ...markFormatting };
};

/** A tracked paragraph property change from `previous`. */
export const paragraphPropertyChange = (
  info: TrackedChangeInfo,
  previous: ParagraphFormatting | undefined,
): ParagraphPropertyChange => {
  const change: ParagraphPropertyChange = { type: "paragraphPropertyChange", info };
  const recorded = paragraphPropertiesOf(previous);
  if (recorded !== undefined) {
    change.previousFormatting = recorded;
  }
  return change;
};

/** The review fields a paragraph states, each only when present. */
export const reviewFieldsOf = (paragraph: Paragraph): ParagraphReviewFields => {
  const fields: ParagraphReviewFields = {};
  if (paragraph.formatting !== undefined) fields.formatting = paragraph.formatting;
  if (paragraph.propertyChanges !== undefined) fields.propertyChanges = paragraph.propertyChanges;
  if (paragraph.pPrMark !== undefined) fields.pPrMark = paragraph.pPrMark;
  return fields;
};

/** The paragraph with exactly these review fields. */
export const withReviewFields = (
  paragraph: Paragraph,
  review: ParagraphReviewFields,
): Paragraph => {
  const next: Paragraph = { ...paragraph };
  delete next.formatting;
  delete next.propertyChanges;
  delete next.pPrMark;
  if (review.formatting !== undefined) next.formatting = review.formatting;
  if (review.propertyChanges !== undefined) next.propertyChanges = review.propertyChanges;
  if (review.pPrMark !== undefined) next.pPrMark = review.pPrMark;
  return next;
};

// ---------------------------------------------------------------------------
// Wrapping leaves in a tracked change
// ---------------------------------------------------------------------------

/** Which tracked change a wrap records. */
export const WRAP_KINDS = Object.freeze({ INSERTION: "insertion", DELETION: "deletion" } as const);

export type WrapKind = (typeof WRAP_KINDS)[keyof typeof WRAP_KINDS];

/** What a wrap does with a record the range runs into but does not cover. */
type SpanningAction = "cut" | "descend" | "keep" | "refuse";

const spanningAction = (node: InlineNode, kind: WrapKind, stamp: RevisionStamp): SpanningAction => {
  if (node.type === "inlineSdt" || node.type === "inlineWrapper") {
    return "descend";
  }
  if (isTrackedWrapper(node) && isAddedRevision(node)) {
    if (kind === WRAP_KINDS.DELETION) return "descend";
    // Another stamp's insertion is split around the new one, never nested in.
    return carriesStamp(node.info, stamp) ? "keep" : "cut";
  }
  if (isRemovedRevisionNode(node)) {
    return kind === WRAP_KINDS.DELETION ? "keep" : "refuse";
  }
  return "cut";
};

/** What a wrap does with a record the range covers whole. */
type CoveredAction = "wrap" | "skip" | "descend" | "refuse";

const coveredAction = (node: InlineNode, kind: WrapKind): CoveredAction => {
  if (isCommentAnchor(node)) {
    return "refuse";
  }
  if (kind === WRAP_KINDS.DELETION) {
    if (isRemovedRevisionNode(node)) return "skip";
    if (isAddedRevision(node)) return "descend";
  }
  return "wrap";
};

type Entry = { node: InlineNode; covered: boolean };

type WrapState = {
  gaps: readonly Gap[];
  kind: WrapKind;
  stamp: RevisionStamp;
  /** The wrappers the wrap created and has not merged away. */
  created: Set<InlineNode>;
  /** Whether the wrap recorded anything: a wrapper created, or merged into a neighbour. */
  recorded: boolean;
  refusal: DocumentOpRefusalReason | undefined;
};

/** Merge a created wrapper into a neighbour carrying the same stamp: the left one, else the right. */
const mergeWithStampedNeighbour = (out: InlineNode[], index: number, state: WrapState): void => {
  const wrapper = out[index];
  if (wrapper === undefined || !isTrackedWrapper(wrapper)) return;
  const matching = (neighbour: InlineNode | undefined): neighbour is TrackedWrapper =>
    neighbour !== undefined &&
    isTrackedWrapper(neighbour) &&
    neighbour.type === wrapper.type &&
    !state.created.has(neighbour) &&
    carriesStamp(neighbour.info, state.stamp);
  const left = out[index - 1];
  if (matching(left)) {
    const merged = mergeLists([left], [wrapper], 1);
    if (merged !== undefined) {
      state.created.delete(wrapper);
      out.splice(index - 1, 2, ...merged);
    }
    return;
  }
  const right = out[index + 1];
  if (matching(right)) {
    // The wrapper that was there keeps its id.
    const merged = mergeLists([wrapper], [right], 1, { mode: "exact", identity: "second" });
    if (merged !== undefined) {
      state.created.delete(wrapper);
      out.splice(index, 2, ...merged);
    }
  }
};

const trackedContent = (
  nodes: readonly InlineNode[],
  state: WrapState,
): TrackedRunContent[] | undefined => {
  const out: TrackedRunContent[] = [];
  for (const node of nodes) {
    if (!isParagraphContent(node) || !isTrackedChild(node)) {
      state.refusal = DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE;
      return undefined;
    }
    out.push(node);
  }
  return out;
};

/** Group the covered entries of a list into wrappers; `undefined` once refused. */
const groupCovered = (entries: readonly Entry[], state: WrapState): InlineNode[] | undefined => {
  const out: InlineNode[] = [];
  const wrapperAt: number[] = [];
  let stretch: InlineNode[] = [];
  const flush = (): boolean => {
    if (stretch.length === 0) return true;
    const content = trackedContent(stretch, state);
    if (content === undefined) return false;
    const info = stampInfo(state.stamp);
    const wrapper: InlineNode =
      state.kind === WRAP_KINDS.INSERTION
        ? { type: "insertion", info, content }
        : { type: "deletion", info, content };
    state.created.add(wrapper);
    state.recorded = true;
    wrapperAt.push(out.length);
    out.push(wrapper);
    stretch = [];
    return true;
  };
  for (const { node, covered } of entries) {
    const action = covered ? coveredAction(node, state.kind) : "skip";
    switch (action) {
      case "refuse":
        state.refusal = DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE;
        return undefined;
      case "skip":
        if (!flush()) return undefined;
        out.push(node);
        break;
      case "descend": {
        if (!flush()) return undefined;
        const children = childNodes(node) ?? [];
        const wrapped = groupCovered(
          children.map((child) => ({ node: child, covered: true })),
          state,
        );
        if (wrapped === undefined) return undefined;
        out.push(rebuildNode(node, wrapped));
        break;
      }
      case "wrap":
        stretch.push(node);
        break;
      default: {
        const unreachable: never = action;
        return unreachable;
      }
    }
  }
  if (!flush()) return undefined;
  for (const index of wrapperAt.toReversed()) mergeWithStampedNeighbour(out, index, state);
  return out;
};

const wrapList = (
  nodes: readonly InlineNode[],
  cursor: Cursor,
  state: WrapState,
): InlineNode[] | undefined => {
  const entries: Entry[] = [];
  for (const node of nodes) {
    const start: Cursor = { ...cursor };
    const measured = partitionNode(node, state.gaps, cursor);
    if (measured.min === measured.max) {
      entries.push({ node, covered: measured.min === 1 });
      continue;
    }
    const action = spanningAction(node, state.kind, state.stamp);
    switch (action) {
      case "refuse":
        state.refusal = DOCUMENT_OP_REFUSAL_REASONS.INSIDE_TRACKED_DELETION;
        return undefined;
      case "keep":
        entries.push({ node, covered: false });
        break;
      case "descend": {
        const end: Cursor = { ...cursor };
        cursor.position = start.position;
        cursor.zeroWidthSeen = start.zeroWidthSeen;
        const children = wrapList(childNodes(node) ?? [], cursor, state);
        if (children === undefined) return undefined;
        cursor.position = end.position;
        cursor.zeroWidthSeen = end.zeroWidthSeen;
        entries.push({ node: rebuildNode(node, children), covered: false });
        break;
      }
      case "cut":
        for (const [region, piece] of measured.pieces) {
          // An insertion does not nest inside another stamp's: that piece gives up its wrapper.
          const unwrap =
            region === 1 && state.kind === WRAP_KINDS.INSERTION && isTrackedWrapper(piece);
          for (const part of unwrap ? (childNodes(piece) ?? []) : [piece]) {
            entries.push({ node: part, covered: region === 1 });
          }
        }
        break;
      default: {
        const unreachable: never = action;
        return unreachable;
      }
    }
  }
  return groupCovered(entries, state);
};

export type WrapOptions = {
  items: readonly ParagraphContent[];
  from: Gap;
  to: Gap;
  kind: WrapKind;
  stamp: RevisionStamp;
};

export type WrapOutcome =
  | { kind: "wrapped"; content: ParagraphContent[] }
  | { kind: "unchanged" }
  | { kind: "refused"; reason: DocumentOpRefusalReason };

/**
 * Wrap the leaves between two gaps in a tracked change carrying the stamp.
 * `unchanged` when there is nothing to record: every leaf is already
 * deleted, or an insertion sits inside a wrapper carrying the same stamp.
 */
export const wrapTracked = ({ items, from, to, kind, stamp }: WrapOptions): WrapOutcome => {
  const state: WrapState = {
    gaps: [from, to],
    kind,
    stamp,
    created: new Set(),
    recorded: false,
    refusal: undefined,
  };
  const content = wrapList(items, startCursor(), state);
  if (content === undefined) {
    return {
      kind: "refused",
      reason: state.refusal ?? DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    };
  }
  if (!state.recorded) {
    return { kind: "unchanged" };
  }
  const out: ParagraphContent[] = [];
  for (const node of content) {
    if (!isParagraphContent(node)) {
      return { kind: "refused", reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE };
    }
    out.push(node);
  }
  return { kind: "wrapped", content: out };
};
