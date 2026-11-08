/**
 * Consecutive tracked edits by one author read as one change: text typed
 * beside the author's own pending insertion goes into it, and a deletion
 * beside the author's own pending deletion takes its stamp, so the new
 * struck piece merges into it. Beside means a text leaf just before or just
 * after the edit, the way the editor finds the change an edit continues;
 * the date or session of the earlier edit does not matter.
 */
import type { ParagraphContent } from "../model/document";
import {
  asParagraphContent,
  compareGaps,
  type Gap,
  type InlineNode,
  type LeafSpan,
  leafSpans,
  rebuildNode,
} from "./leaves";
import { carriesStamp, isTrackedWrapper, type TrackedWrapper } from "./review";
import type { InlineSlice, RevisionStamp } from "./types";

/** The leaf just before a gap and the leaf just after it, partly covered ones included. */
const leavesBeside = (
  spans: readonly LeafSpan[],
  gap: Gap,
): readonly [LeafSpan | undefined, LeafSpan | undefined] => [
  spans.findLast(({ before }) => compareGaps(before, gap) < 0),
  spans.find(({ after }) => compareGaps(after, gap) > 0),
];

/** The innermost tracked record of `kind` holding a text leaf, when `author` made it. */
const ownRecord = (
  span: LeafSpan | undefined,
  kind: "insertion" | "deletion",
  author: string,
): TrackedWrapper | undefined => {
  if (span?.node.type !== "text") return undefined;
  const record = span.ancestors.findLast(
    (ancestor) => isTrackedWrapper(ancestor) && ancestor.type === kind,
  );
  return record !== undefined &&
    isTrackedWrapper(record) &&
    record.info.author === author &&
    record.info.suggestionId === undefined
    ? record
    : undefined;
};

type GroupedDeletionOptions = {
  /** The paragraph content at the range's start, and the start's gap. */
  start: { items: readonly ParagraphContent[]; gap: Gap };
  /** The paragraph content at the range's end, and the end's gap. */
  end: { items: readonly ParagraphContent[]; gap: Gap };
  stamp: RevisionStamp;
};

/**
 * The stamp a tracked deletion of the range takes: that of the author's own
 * pending deletion beside the range, looked for before and after its start,
 * then before and after its end, with the deletion's own id; else `stamp`.
 */
export const groupedDeletionStamp = ({
  start,
  end,
  stamp,
}: GroupedDeletionOptions): RevisionStamp => {
  if (stamp.suggestionId !== undefined) return stamp;
  const beside = [
    ...leavesBeside(leafSpans(start.items), start.gap),
    ...leavesBeside(leafSpans(end.items), end.gap),
  ];
  for (const span of beside) {
    const record = ownRecord(span, "deletion", stamp.author);
    const date = record?.info.date;
    if (record === undefined || date === undefined) continue;
    const continued: RevisionStamp = {
      id: stamp.id,
      author: record.info.author,
      date,
      ...(record.info.initials === undefined
        ? {}
        : { initials: record.info.initials }),
    };
    // A record whose metadata no stamp writes stays a change of its own.
    if (carriesStamp(record.info, continued)) return continued;
  }
  return stamp;
};

/**
 * Where text typed at a caret goes when it continues the author's own
 * pending insertion beside it, the one before the caret first: `content`
 * inside that insertion and the containers the typed text continues anyway,
 * open into them on each side the source holds them. `undefined` when no
 * such insertion is beside the caret.
 */
export const ownInsertionSlice = ({
  items,
  gap,
  continued,
  author,
  content,
}: {
  items: readonly ParagraphContent[];
  gap: Gap;
  /** The content controls and wrappers the typed text continues, outermost first. */
  continued: readonly InlineNode[];
  author: string;
  content: readonly ParagraphContent[];
}): InlineSlice | undefined => {
  const [before, after] = leavesBeside(leafSpans(items), gap);
  for (const span of [before, after]) {
    if (
      span === undefined ||
      !continued.every(
        (container, depth) => span.ancestors[depth] === container,
      )
    )
      continue;
    const record = ownRecord(span, "insertion", author);
    // Only directly inside what the typed text continues: an insertion
    // inside a hyperlink it reaches the edge of stays where it is.
    if (record === undefined || span.ancestors[continued.length] !== record)
      continue;
    const chain = [...continued, record];
    const open = (side: LeafSpan | undefined) => {
      let depth = 0;
      while (depth < chain.length && side?.ancestors[depth] === chain[depth])
        depth += 1;
      return depth;
    };
    let wrapped: InlineNode[] = content.slice();
    for (const container of chain.toReversed())
      wrapped = [rebuildNode(container, wrapped)];
    return {
      openStart: open(before),
      openEnd: open(after),
      content: asParagraphContent(wrapped),
    };
  }
  return undefined;
};
