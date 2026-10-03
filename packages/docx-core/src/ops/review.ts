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
  spanningRecords,
} from "./leaves";
import { isTrackedChild } from "./offsets";
import { identitySlots } from "./slots";
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
export const isAddedRevision = (
  node: InlineNode,
): node is Extract<TrackedWrapper, { type: "insertion" | "moveTo" }> =>
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
): boolean => structurallyEqual(paragraphPropertiesOf(left), paragraphPropertiesOf(right));

/** Whether two property sets state the same run properties for the paragraph mark. */
export const sameMarkFormatting = (
  left: ParagraphFormatting | undefined,
  right: ParagraphFormatting | undefined,
): boolean => structurallyEqual(markFormattingOf(left) ?? {}, markFormattingOf(right) ?? {});

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

type FoldParagraphPropertyChangeOptions = {
  paragraph: Paragraph;
  formatting: ParagraphFormatting | undefined;
  stamp: RevisionStamp;
};

/** Fold paragraph formatting into one review, retaining its id and original baseline. */
export const foldedParagraphPropertyChange = ({
  paragraph,
  formatting,
  stamp,
}: FoldParagraphPropertyChangeOptions): ParagraphPropertyChange => {
  const existing = paragraph.propertyChanges?.at(0);
  if (existing === undefined)
    return paragraphPropertyChange(stampInfo(stamp), paragraph.formatting);
  const info = { ...existing.info, ...stampInfo(stamp), id: existing.info.id };
  if (stamp.initials === undefined) delete info.initials;
  // A companion date belongs to the previous attribution, not the new stamp.
  delete info.utcDate;
  const change = { ...existing, info };
  // Parsed reviews carry this derived convenience field; keep it tied to live properties.
  if (existing.currentFormatting !== undefined) {
    if (formatting === undefined) delete change.currentFormatting;
    else change.currentFormatting = formatting;
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

type Entry = {
  node: InlineNode;
  covered: boolean;
  joinBefore?: number;
  joinAfter?: number;
  retainedAfter?: NonNullable<TrackedWrapper["resolutionJoins"]>["retainedAfter"];
};

type WrapState = {
  acceptance: NonNullable<TrackedWrapper["resolutionJoins"]>["acceptance"];
  gaps: readonly [Gap, Gap];
  resolutionJoin: number | undefined;
  depth: number;
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
    wrapper.resolutionJoins?.acceptance === undefined &&
    neighbour.resolutionJoins?.acceptance === undefined &&
    !(
      wrapper.resolutionJoins !== undefined &&
      neighbour.resolutionJoins !== undefined &&
      (wrapper.resolutionJoins.before > 0 ||
        wrapper.resolutionJoins.after > 0 ||
        wrapper.resolutionJoins.remove > 0 ||
        neighbour.resolutionJoins.before > 0 ||
        neighbour.resolutionJoins.after > 0 ||
        neighbour.resolutionJoins.remove > 0)
    ) &&
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
  let stretch: Entry[] = [];
  const flush = (): boolean => {
    if (stretch.length === 0) return true;
    const content = trackedContent(
      stretch.map(({ node }) => node),
      state,
    );
    if (content === undefined) return false;
    const info = stampInfo(state.stamp);
    const wrapper: InlineNode =
      state.kind === WRAP_KINDS.INSERTION
        ? { type: "insertion", info, content }
        : { type: "deletion", info, content };
    if (state.resolutionJoin !== undefined && isTrackedWrapper(wrapper)) {
      wrapper.resolutionJoins = {
        before: stretch.at(0)?.joinBefore ?? 0,
        after: stretch.at(-1)?.joinAfter ?? 0,
        remove: Math.max(0, state.resolutionJoin - state.depth),
      };
      if (state.acceptance !== undefined) wrapper.resolutionJoins.acceptance = state.acceptance;
      const retainedAfter = stretch.at(-1)?.retainedAfter;
      if (retainedAfter !== undefined) wrapper.resolutionJoins.retainedAfter = retainedAfter;
    }
    state.created.add(wrapper);
    state.recorded = true;
    wrapperAt.push(out.length);
    out.push(wrapper);
    stretch = [];
    return true;
  };
  for (const entry of entries) {
    const { node, covered } = entry;
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
        state.depth += 1;
        const wrapped = groupCovered(
          children.map((child) => ({ node: child, covered: true })),
          state,
        );
        state.depth -= 1;
        if (wrapped === undefined) return undefined;
        out.push(rebuildNode(node, wrapped));
        break;
      }
      case "wrap":
        stretch.push(entry);
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
        state.depth += 1;
        const children = wrapList(childNodes(node) ?? [], cursor, state);
        state.depth -= 1;
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
            const localGap = (gap: Gap): Gap => ({
              offset: gap.offset - start.position,
              zeroWidthBefore:
                gap.offset === start.position
                  ? gap.zeroWidthBefore - start.zeroWidthSeen
                  : gap.zeroWidthBefore,
            });
            const beforeChain = spanningRecords([node], [localGap(state.gaps[0])], 0, 1);
            const afterChain = spanningRecords([node], [localGap(state.gaps[1])], 0, 1);
            const retainedAfter =
              region === 1 && state.kind === WRAP_KINDS.DELETION
                ? afterChain.flatMap((record, depth) => {
                    const source = beforeChain.includes(record) ? [] : identitySlots(record);
                    return source.length === 0 ? [] : [{ depth, source, target: source }];
                  })
                : [];
            entries.push({
              node: part,
              ...(retainedAfter.length === 0 ? {} : { retainedAfter }),
              covered: region === 1,
              joinBefore:
                region === 1 ? spanningRecords([node], [localGap(state.gaps[0])], 0, 1).length : 0,
              joinAfter:
                region === 1 ? spanningRecords([node], [localGap(state.gaps[1])], 0, 1).length : 0,
            });
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
  acceptance?: NonNullable<TrackedWrapper["resolutionJoins"]>["acceptance"];
  items: readonly ParagraphContent[];
  from: Gap;
  to: Gap;
  kind: WrapKind;
  stamp: RevisionStamp;
  /** Exact top-level cut depth restored when this wrapper content is removed. */
  resolutionJoin?: number;
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
export const wrapTracked = ({
  items,
  from,
  to,
  kind,
  stamp,
  resolutionJoin,
  acceptance,
}: WrapOptions): WrapOutcome => {
  const state: WrapState = {
    gaps: [from, to],
    resolutionJoin,
    acceptance,
    depth: 0,
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
