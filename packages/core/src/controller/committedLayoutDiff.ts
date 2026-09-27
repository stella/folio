// What an incremental layout pass may re-measure. The committed layout's
// measures were computed for one document (and one template preview); the pass
// lays out another. The range worth re-measuring is where the two differ, so it
// is derived from the two, never carried alongside a scheduled pass: a carried
// range is relative to whatever document the caller had in mind, which is how a
// pass reused one document's measures for another document's blocks (#1142).

import type { Node as PMNode } from "prosemirror-model";

import {
  templatePreviewDirtyRange,
  type TemplatePreviewFlowState,
} from "../layout-bridge/convert/templatePreviewFlow";
import type { DirtyRange } from "../paged-layout/incrementalMeasure";
import type { LayoutTemplatePreview } from "./layoutSession";

export type LayoutFlowSource = {
  doc: PMNode;
  preview: LayoutTemplatePreview;
};

export type CommittedLayoutDiff =
  | { type: "full" }
  | {
      type: "range";
      /** In the next document's positions. */
      range: DirtyRange;
    };

const FULL: CommittedLayoutDiff = { type: "full" };

/** Where two documents differ: `[start, endBefore)` became `[start, endAfter)`. */
type DocumentDiff = {
  start: number;
  endBefore: number;
  endAfter: number;
};

const diffDocuments = (before: PMNode, after: PMNode): DocumentDiff | null => {
  const start = before.content.findDiffStart(after.content);
  if (start === null) {
    return null;
  }
  const end = before.content.findDiffEnd(after.content);
  if (end === null) {
    return null;
  }
  // Repeated content lets the scans from both ends cross; widen both ends by
  // the overlap so the changed span starts at `start` in both documents.
  const overlap = start - Math.min(end.a, end.b);
  return overlap > 0
    ? { start, endBefore: end.a + overlap, endAfter: end.b + overlap }
    : { start, endBefore: end.a, endAfter: end.b };
};

/** Map a range of the previous document into the next one's positions. */
const mapRangeThroughDiff = <T extends { from: number; to: number }>(
  range: T,
  diff: DocumentDiff,
): T => {
  const shift = diff.endAfter - diff.endBefore;
  const mapPosition = (pos: number, insideChange: number): number => {
    if (pos < diff.start) {
      return pos;
    }
    if (pos >= diff.endBefore) {
      return pos + shift;
    }
    return insideChange;
  };
  return {
    ...range,
    from: mapPosition(range.from, diff.start),
    to: mapPosition(range.to, diff.endAfter),
  };
};

const mapPreviewThroughDiff = (
  preview: LayoutTemplatePreview,
  diff: DocumentDiff,
): TemplatePreviewFlowState => ({
  entries: preview.entries.map((entry) => mapRangeThroughDiff(entry, diff)),
  hidden: preview.hidden.map((range) => mapRangeThroughDiff(range, diff)),
});

const unionRanges = (first: DirtyRange | null, second: DirtyRange | null): DirtyRange | null => {
  if (!first) {
    return second;
  }
  if (!second) {
    return first;
  }
  return { from: Math.min(first.from, second.from), to: Math.max(first.to, second.to) };
};

/**
 * The span of `next` whose flow blocks may differ from the ones the committed
 * layout measured for `committed`, or `full` when no span bounds it.
 *
 * `full` also covers "nothing in the body flow changed": such a pass was
 * requested for an input outside it (a header, footer or note story).
 */
export const diffAgainstCommittedLayout = (
  committed: LayoutFlowSource,
  next: LayoutFlowSource,
): CommittedLayoutDiff => {
  // A preview mode switch restyles every substituted run.
  if (committed.preview.mode !== next.preview.mode) {
    return FULL;
  }
  // Document-level attributes (the final section's start, say) reach every block.
  if (!committed.doc.sameMarkup(next.doc)) {
    return FULL;
  }
  const docDiff = diffDocuments(committed.doc, next.doc);
  let range = docDiff ? { from: docDiff.start, to: docDiff.endAfter } : null;
  const previewChanged =
    committed.preview.entries !== next.preview.entries ||
    committed.preview.hidden !== next.preview.hidden;
  if (previewChanged) {
    const previousPreview = docDiff
      ? mapPreviewThroughDiff(committed.preview, docDiff)
      : committed.preview;
    range = unionRanges(range, templatePreviewDirtyRange(previousPreview, next.preview));
  }
  return range ? { type: "range", range } : FULL;
};
