import type { Node as PMNode } from "prosemirror-model";
import { StepMap } from "prosemirror-transform";

import type { RevisionDecision } from "@stll/docx-core/ops";

import { COMMENT_REFERENCE_NODE_NAME } from "./commentReferenceIntegrity";
import { BLOCK_MOVE_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/BlockMoveRangeBoundaryExtension";
import { MOVE_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/MoveRangeBoundaryExtension";
import { RANGE_ANCHOR_NODE_NAME } from "./extensions/nodes/RangeAnchorExtension";
import { isUsableParaId } from "./paraIds";
import { trackedRevisionPathOf } from "./trackedRevisionPath";

type Span = {
  start: number;
  width: number;
  key: string;
  text?: string;
  provenance?: string;
  paragraphId: string | undefined;
  /** A tie-breaker for repeated content. The output decides what survived. */
  pending: boolean;
};

type ParagraphEnds = Map<string, number | null>;
type StorySpans = { spans: Span[]; ends: ParagraphEnds };
const NO_REVISION_LAYERS = [] as const;

const opposingRevision = (decision: RevisionDecision) =>
  decision === "accept" ? "deletion" : "insertion";

const atomKey = (node: PMNode): string => {
  const marker: unknown = node.attrs["marker"];
  if (
    typeof marker === "object" &&
    marker !== null &&
    "type" in marker &&
    typeof marker.type === "string" &&
    "id" in marker &&
    typeof marker.id === "number"
  )
    return `leaf:${node.type.name}:${marker.type}:${marker.id}`;
  const id =
    node.attrs["commentId"] ??
    node.attrs["noteId"] ??
    node.attrs["imageId"] ??
    node.attrs["id"];
  return typeof id === "string" || typeof id === "number"
    ? `leaf:${node.type.name}:${id}`
    : `leaf:${node.type.name}`;
};

const storySpans = (doc: PMNode, decision?: RevisionDecision): StorySpans => {
  const spans: Span[] = [];
  const ends: ParagraphEnds = new Map();
  const opposed =
    decision === undefined ? undefined : opposingRevision(decision);
  const visit = (
    node: PMNode,
    start: number,
    parent: { pending: boolean; paragraphId: string | undefined },
  ): void => {
    const name = node.type.name;
    const paraId = name === "paragraph" ? node.attrs["paraId"] : undefined;
    const paragraphId = isUsableParaId(paraId)
      ? paraId.toUpperCase()
      : parent.paragraphId;
    const cellMarker: unknown = node.attrs["cellMarker"];
    const paths =
      node.marks.length === 0
        ? NO_REVISION_LAYERS
        : node.marks.flatMap((mark) =>
            mark.type.name === "insertion" || mark.type.name === "deletion"
              ? trackedRevisionPathOf(mark)
              : NO_REVISION_LAYERS,
          );
    const provenance =
      paths.length === 0
        ? {}
        : {
            provenance: paths
              .map(({ type, revisionId }) => `${type}:${revisionId}`)
              .join("/"),
          };
    const pending =
      parent.pending ||
      (opposed !== undefined &&
        (name === MOVE_RANGE_BOUNDARY_NODE_NAME ||
          name === BLOCK_MOVE_RANGE_BOUNDARY_NODE_NAME ||
          (name === "tableRow" &&
            node.attrs[opposed === "insertion" ? "trIns" : "trDel"] != null) ||
          (typeof cellMarker === "object" &&
            cellMarker !== null &&
            "kind" in cellMarker &&
            cellMarker.kind === (opposed === "insertion" ? "ins" : "del")) ||
          paths.some(({ type }) =>
            opposed === "deletion"
              ? type === "deletion" || type === "moveFrom"
              : type === "insertion" || type === "moveTo",
          )));
    if (node.text !== undefined) {
      spans.push({
        start,
        width: node.text.length,
        key: "text",
        text: node.text,
        pending,
        paragraphId,
        ...provenance,
      });
      return;
    }
    if (node.isLeaf) {
      spans.push({
        start,
        width: 1,
        key: atomKey(node),
        pending,
        paragraphId,
        ...provenance,
      });
      return;
    }
    spans.push({ start, width: 1, key: `open:${name}`, pending, paragraphId });
    const scope = { pending, paragraphId };
    node.forEach((child, offset) => visit(child, start + offset + 1, scope));
    if (paragraphId !== undefined && isUsableParaId(paraId))
      ends.set(paragraphId, ends.has(paragraphId) ? null : spans.length);
    spans.push({
      start: start + node.nodeSize - 1,
      width: 1,
      key: `close:${name}`,
      pending,
      paragraphId,
    });
  };
  const root = { pending: false, paragraphId: undefined };
  doc.forEach((child, start) => visit(child, start, root));
  return { spans, ends };
};

type ContiguousMatch = { before: number; after: number; width: number };
type Match = (span: Readonly<ContiguousMatch>) => void;

type Region = {
  before: readonly Span[];
  after: readonly Span[];
};

/**
 * Match by span, across text node seams. A pending span can disappear unless
 * the result still names its provenance. Keep the candidate only when it
 * accounts for the entire resolved region.
 */
const matchStream = (
  { before, after }: Region,
  skipPending: boolean,
): ContiguousMatch[] | undefined => {
  const matches: ContiguousMatch[] = [];
  let leftIndex = 0;
  let rightIndex = 0;
  let leftOffset = 0;
  let rightOffset = 0;
  for (;;) {
    const right = after[rightIndex];
    let left = before[leftIndex];
    while (
      skipPending &&
      leftOffset === 0 &&
      left?.pending === true &&
      (left.provenance === undefined || left.provenance !== right?.provenance)
    )
      left = before[++leftIndex];
    if (!left || !right) return left || right ? undefined : matches;
    const width = Math.min(left.width - leftOffset, right.width - rightOffset);
    if (
      left.key !== right.key ||
      (right.provenance !== undefined &&
        left.provenance !== right.provenance) ||
      left.text?.slice(leftOffset, leftOffset + width) !==
        right.text?.slice(rightOffset, rightOffset + width)
    )
      return undefined;
    const leftStart = left.start + leftOffset;
    const rightStart = right.start + rightOffset;
    const previous = matches.at(-1);
    if (
      previous &&
      previous.before + previous.width === leftStart &&
      previous.after + previous.width === rightStart
    )
      previous.width += width;
    else matches.push({ before: leftStart, after: rightStart, width });
    leftOffset += width;
    rightOffset += width;
    if (leftOffset === left.width) {
      leftIndex++;
      leftOffset = 0;
    }
    if (rightOffset === right.width) {
      rightIndex++;
      rightOffset = 0;
    }
  }
};

type Unit = {
  position: number;
  key: string;
  character: string | undefined;
  provenance: string | undefined;
  paragraphId: string | undefined;
  pending: boolean;
};

const unitsOf = (spans: readonly Span[]): Unit[] => {
  const units: Unit[] = [];
  for (const {
    start,
    width,
    key,
    text,
    pending,
    provenance,
    paragraphId,
  } of spans)
    for (let offset = 0; offset < width; offset++)
      units.push({
        position: start + offset,
        key,
        character: text?.charAt(offset),
        pending,
        provenance,
        paragraphId,
      });
  return units;
};

const sameUnit = (left: Unit, right: Unit): boolean =>
  left.key === right.key && left.character === right.character;

/** Limit ambiguous matching to a small region between surviving paragraph IDs. */
const MAX_ALIGNMENT_CELLS = 65_536;
const MAX_ALIGNMENT_EDGE = 2_048;

const sameStandingSpan = (
  left: Span | undefined,
  right: Span | undefined,
): boolean =>
  left !== undefined &&
  right !== undefined &&
  left.key === right.key &&
  left.width === right.width &&
  left.text === right.text &&
  left.provenance === right.provenance;

/** Keep matching edge spans when the interior exceeds the alignment budget. */
const matchStandingEdges = ({ before, after }: Region, match: Match): void => {
  const sameSpan = (left: Span | undefined, right: Span | undefined) =>
    left?.pending === false &&
    sameStandingSpan(left, right) &&
    left.paragraphId === right?.paragraphId;
  let start = 0;
  for (const span of before) {
    const target = after[start];
    if (!target || !sameSpan(span, target)) break;
    match({ before: span.start, after: target.start, width: span.width });
    start++;
  }
  let oldEnd = before.length;
  let newEnd = after.length;
  while (
    oldEnd > start &&
    newEnd > start &&
    sameSpan(before[oldEnd - 1], after[newEnd - 1])
  ) {
    oldEnd--;
    newEnd--;
  }
  for (const span of before.slice(oldEnd)) {
    const target = after[newEnd++];
    if (target)
      match({ before: span.start, after: target.start, width: span.width });
  }
};

const isAddressEdge = (span: Span): boolean =>
  span.key.startsWith("open:") ||
  span.key.startsWith("close:") ||
  span.key.startsWith(`leaf:${COMMENT_REFERENCE_NODE_NAME}:`) ||
  span.key.startsWith(`leaf:${MOVE_RANGE_BOUNDARY_NODE_NAME}`) ||
  span.key.startsWith(`leaf:${BLOCK_MOVE_RANGE_BOUNDARY_NODE_NAME}`) ||
  span.key.startsWith(`leaf:${RANGE_ANCHOR_NODE_NAME}`);

/** The addresses a forward scan passes before its next content match. */
const addressGap = (spans: readonly Span[]) => {
  let index = 0;
  return (to: number): Span[] => {
    const edges: Span[] = [];
    for (;;) {
      const span = spans[index];
      if (span === undefined || span.start >= to) return edges;
      index++;
      if (isAddressEdge(span)) edges.push(span);
    }
  };
};

/** Resolution can join long runs while retiring orphaned comment references. */
const matchContent = ({ before, after }: Region, match: Match): boolean => {
  const content = {
    before: before.filter((span) => !isAddressEdge(span)),
    after: after.filter((span) => !isAddressEdge(span)),
  };
  if (
    content.before.length === before.length &&
    content.after.length === after.length
  )
    return false;
  const aligned = matchStream(content, true) ?? matchStream(content, false);
  if (aligned === undefined || aligned.length === 0) return false;
  const leftGap = addressGap(before);
  const rightGap = addressGap(after);
  for (const span of aligned) {
    // A surviving opening edge can sit behind discarded table wrappers
    // and ahead of a synthesized range marker in the same gap.
    matchBounded(
      { before: leftGap(span.before), after: rightGap(span.after) },
      match,
    );
    match(span);
  }
  matchBounded(
    {
      before: leftGap(Number.POSITIVE_INFINITY),
      after: rightGap(Number.POSITIVE_INFINITY),
    },
    match,
  );
  return true;
};

/** Maximize retained positions, favouring ordinary text when repeats tie. */
const matchBounded = ({ before, after }: Region, match: Match): void => {
  const beforeWidth = before.reduce((total, span) => total + span.width, 0);
  const afterWidth = after.reduce((total, span) => total + span.width, 0);
  if (
    beforeWidth === 0 ||
    afterWidth === 0 ||
    Math.max(beforeWidth, afterWidth) > MAX_ALIGNMENT_EDGE ||
    beforeWidth * afterWidth > MAX_ALIGNMENT_CELLS
  ) {
    matchStandingEdges({ before, after }, match);
    return;
  }
  const left = unitsOf(before);
  const right = unitsOf(after);
  const rowSize = right.length + 1;
  const scores = new Float64Array((left.length + 1) * rowSize);
  const paragraphWeight = Math.min(left.length, right.length) + 1;
  const provenanceWeight = paragraphWeight * paragraphWeight;
  const matchWeight = provenanceWeight * paragraphWeight;
  const weight = (l: Unit, r: Unit): number =>
    matchWeight +
    Number(!l.pending) +
    (r.paragraphId !== undefined && l.paragraphId === r.paragraphId
      ? paragraphWeight
      : 0) +
    (r.provenance !== undefined && l.provenance === r.provenance
      ? provenanceWeight
      : 0);
  const reversedRight = right.toReversed();
  for (const [offset, l] of left.toReversed().entries()) {
    const i = left.length - 1 - offset;
    for (const [rightOffset, r] of reversedRight.entries()) {
      const j = right.length - 1 - rightOffset;
      const at = i * rowSize + j;
      scores[at] = Math.max(
        scores[at + rowSize] ?? 0,
        scores[at + 1] ?? 0,
        sameUnit(l, r) ? (scores[at + rowSize + 1] ?? 0) + weight(l, r) : 0,
      );
    }
  }
  let i = 0;
  let j = 0;
  const span = { before: 0, after: 0, width: 1 };
  for (;;) {
    const l = left[i];
    const r = right[j];
    if (!l || !r) return;
    const at = i * rowSize + j;
    if (
      sameUnit(l, r) &&
      scores[at] === (scores[at + rowSize + 1] ?? 0) + weight(l, r)
    ) {
      span.before = l.position;
      span.after = r.position;
      match(span);
      i++;
      j++;
    } else if (scores[at] === scores[at + rowSize]) i++;
    else j++;
  }
};

const matchRegion = (region: Region, match: Match): void => {
  const aligned = matchStream(region, true) ?? matchStream(region, false);
  if (aligned === undefined) {
    if (!matchContent(region, match)) matchBounded(region, match);
    return;
  }
  for (const span of aligned) match(span);
};

type MarkupViewPositionMapOptions = {
  before: PMNode;
  after: PMNode;
  decision: RevisionDecision;
};

/**
 * Match positions between authored and resolved projections. A surviving
 * paragraph's closing edge anchors the region: on a join, preceding content
 * can move into that paragraph, before its original opening edge. PM
 * positions outside monotone matches are replacement ranges.
 */
export const markupViewPositionMap = ({
  before,
  after,
  decision,
}: MarkupViewPositionMapOptions): StepMap => {
  if (before === after) return StepMap.empty;
  const left = storySpans(before, decision);
  const right = storySpans(after);
  const ranges: number[] = [];
  let oldEnd = 0;
  let newEnd = 0;
  const match: Match = ({ before: oldStart, after: newStart, width }) => {
    if (oldStart !== oldEnd || newStart !== newEnd)
      ranges.push(oldEnd, oldStart - oldEnd, newStart - newEnd);
    oldEnd = oldStart + width;
    newEnd = newStart + width;
  };
  let oldIndex = 0;
  let newIndex = 0;
  const anchor = { before: 0, after: 0, width: 1 };
  for (const [id, leftIndex] of left.ends) {
    const rightIndex = right.ends.get(id);
    if (
      leftIndex == null ||
      rightIndex == null ||
      leftIndex < oldIndex ||
      rightIndex < newIndex
    )
      continue;
    const l = left.spans[leftIndex];
    const r = right.spans[rightIndex];
    if (!l || !r) continue;
    matchRegion(
      {
        before: left.spans.slice(oldIndex, leftIndex),
        after: right.spans.slice(newIndex, rightIndex),
      },
      match,
    );
    anchor.before = l.start;
    anchor.after = r.start;
    match(anchor);
    oldIndex = leftIndex + 1;
    newIndex = rightIndex + 1;
  }
  matchRegion(
    { before: left.spans.slice(oldIndex), after: right.spans.slice(newIndex) },
    match,
  );
  match({ before: before.content.size, after: after.content.size, width: 0 });
  return ranges.length > 0 ? new StepMap(ranges) : StepMap.empty;
};
