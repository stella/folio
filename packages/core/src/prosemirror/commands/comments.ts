/** Comment mark commands and review queries over an editor projection. */
import { type Mark, type Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";

import { sameStatedParagraphNumbering } from "../../docx/numberingReference";
import { expectParagraphAttrs } from "../attrs";
import { enclosingRevisionIds } from "../contentControlRevisions";
import { INLINE_CONTENT_CONTROL_NODE_NAME } from "../extensions/nodes/SdtExtension";
import { isPPrMarkAttr, paragraphMarkWasAdded } from "../revisionCarriers";
import type { ParagraphPropertyChangeAttrs } from "../schema/nodes";
import {
  extractTrackedChanges,
  type TrackedChangeEntry,
} from "../utils/extractTrackedChanges";

export { isPPrMarkAttr, paragraphMarkWasAdded } from "../revisionCarriers";
export {
  findAIEditRevisionRange,
  isTableCellRevisionAttr,
  isTableRowRevisionAttr,
  type TableRowRevisionAttr,
} from "../revisionRangeQueries";
export {
  collectSuggestions,
  findSuggestionRange,
  getSuggestions,
  isSuggestionMark,
  mergeSegments,
  pendingSuggestionRevisionIds,
  readStructuralSuggestion,
  readSuggestedMarker,
  readSuggestedParagraphMark,
  readSuggestedParagraphPropertyChange,
  suggestedCommentOf,
  suggestionCommentIds,
  suggestionIdOfMarkRecord,
  suggestionIdOfRevision,
  suggestionOwnersByRevision,
  type FolioSuggestion,
  type SuggestionAppliedAs,
  type SuggestionKind,
} from "../suggestionQueries";

/**
 * Add a comment mark to the current selection.
 */
export function addCommentMark(commentId: number): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty) {
      return false;
    }

    const commentType = state.schema.marks["comment"];
    if (!commentType) {
      return false;
    }

    if (dispatch) {
      const tr = state.tr.addMark(from, to, commentType.create({ commentId }));
      dispatch(tr);
    }
    return true;
  };
}

/**
 * Remove a comment mark by ID from the entire document.
 */
export function removeCommentMark(commentId: number): Command {
  return (state, dispatch) => {
    const commentType = state.schema.marks["comment"];
    if (!commentType) {
      return false;
    }

    if (dispatch) {
      const tr = state.tr;
      state.doc.descendants((node, pos) => {
        if (node.isText) {
          for (const mark of node.marks) {
            if (
              mark.type === commentType &&
              mark.attrs["commentId"] === commentId
            ) {
              tr.removeMark(pos, pos + node.nodeSize, mark);
            }
          }
        }
      });
      if (tr.steps.length > 0) {
        dispatch(tr);
      }
    }
    return true;
  };
}

export type ParagraphBoundaryChange = {
  from: number;
  to: number;
  type: "insertion" | "deletion";
  author?: string;
  date?: string;
  revisionId?: number;
};

type RevisionInfoAttrs = {
  id?: unknown;
  author?: unknown;
  date?: unknown;
};

function readRevisionInfo(info: RevisionInfoAttrs | undefined): {
  author?: string;
  date?: string;
  revisionId?: number;
} {
  const revision: { author?: string; date?: string; revisionId?: number } = {};
  if (typeof info?.author === "string") {
    revision.author = info.author;
  }
  if (typeof info?.date === "string") {
    revision.date = info.date;
  }
  if (typeof info?.id === "number") {
    revision.revisionId = info.id;
  }
  return revision;
}

function getListPropertyChangeType(
  node: PMNode,
  change: ParagraphPropertyChangeAttrs,
): ParagraphBoundaryChange["type"] | null {
  const previousFormatting = change.previousFormatting;
  if (
    previousFormatting == null ||
    !Object.hasOwn(previousFormatting, "numPr")
  ) {
    return null;
  }

  // `null` is the stored record's tombstone and `undefined` the attr's
  // absence; stated equality over the rest has one owner.
  const current = expectParagraphAttrs(node).numPr;
  const previous = previousFormatting["numPr"] ?? undefined;
  if (sameStatedParagraphNumbering(previous, current)) {
    return null;
  }
  return current === undefined ? "deletion" : "insertion";
}

function toParagraphBoundaryChange(
  node: PMNode,
  pos: number,
  type: ParagraphBoundaryChange["type"],
  info?: RevisionInfoAttrs,
): ParagraphBoundaryChange {
  return {
    from: pos + node.nodeSize - 1,
    to: pos + node.nodeSize,
    type,
    ...readRevisionInfo(info),
  };
}

export function findParagraphBoundaryChangeAtPosition(
  state: EditorState,
  pos: number,
): ParagraphBoundaryChange | null {
  const $pos = state.doc.resolve(pos);
  const node = $pos.parent;
  if (node.type.name !== "paragraph") {
    return null;
  }

  const paragraphPos = $pos.before($pos.depth);
  const pPrMark = node.attrs["pPrMark"];
  if (isPPrMarkAttr(pPrMark)) {
    return toParagraphBoundaryChange(
      node,
      paragraphPos,
      // A relocation's destination break was added exactly as an insertion's
      // was; naming the kinds one by one here read `moveTo` as a deletion and
      // showed the reader a paragraph arriving as one going away.
      paragraphMarkWasAdded(pPrMark.kind) ? "insertion" : "deletion",
      pPrMark.info,
    );
  }

  const propertyChanges = expectParagraphAttrs(node)._propertyChanges;
  if (!Array.isArray(propertyChanges)) {
    return null;
  }

  for (const change of propertyChanges) {
    const type = getListPropertyChangeType(node, change);
    if (type) {
      return toParagraphBoundaryChange(node, paragraphPos, type, change.info);
    }
  }

  return null;
}

/** Every revision id one sidebar entry stands for. */
export const trackedChangeEntryRevisionIds = (
  entry: TrackedChangeEntry,
): number[] => [
  ...new Set([
    entry.revisionId,
    ...(entry.insertionRevisionId === undefined
      ? []
      : [entry.insertionRevisionId]),
    ...(entry.coalescedRevisionIds ?? []),
  ]),
];

type ChangeRange = {
  from: number;
  to: number;
  type: "insertion" | "deletion";
};

type PairedReplacementRange = { from: number; to: number };

const pairedReplacementRangesByDoc = new WeakMap<
  PMNode,
  readonly PairedReplacementRange[]
>();

const pairedReplacementRanges = (state: EditorState) => {
  const cached = pairedReplacementRangesByDoc.get(state.doc);
  if (cached) return cached;

  const ranges: PairedReplacementRange[] = [];
  const { entries } = extractTrackedChanges(state);
  for (const entry of entries) {
    if (entry.type !== "replacement" || entry.insertionRevisionId === undefined)
      continue;
    ranges.push({ from: entry.from, to: entry.to });
  }
  pairedReplacementRangesByDoc.set(state.doc, ranges);
  return ranges;
};

type PairedReplacementRangeAtOptions = {
  ranges: readonly PairedReplacementRange[];
  from: number;
  to: number;
};

const pairedReplacementRangeAt = ({
  ranges,
  from,
  to,
}: PairedReplacementRangeAtOptions): PairedReplacementRange | undefined => {
  let low = 0;
  let high = ranges.length;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    if (ranges[middle]!.from <= from) low = middle + 1;
    else high = middle;
  }
  const candidate = ranges[low - 1];
  return candidate && from >= candidate.from && to <= candidate.to
    ? candidate
    : undefined;
};

/**
 * Find the tracked change mark range at a given cursor position.
 * If the cursor is inside a tracked change, returns the full extent
 * of that mark (expanding to cover all adjacent nodes with the same
 * revision ID). If from !== to (range selection), returns {from, to}.
 */
export function findChangeAtPosition(
  state: EditorState,
  from: number,
  to: number,
): { from: number; to: number } {
  // If there's a range selection, use it directly
  if (from !== to) {
    return { from, to };
  }

  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  if (!insertionType && !deletionType) {
    return { from, to };
  }

  // Resolve the position and check marks at cursor. The parent may be an
  // inline container, such as a content control, holding the text.
  const $pos = state.doc.resolve(from);
  const node = $pos.parent;
  if (!node.inlineContent) {
    return { from, to };
  }

  // Find the text node at this position and its mark instance. We capture
  // the specific instance (not just the type) so the adjacency expansion
  // below stays inside a single revision — two back-to-back insertions
  // belonging to different `revisionId`s must not be treated as one range.
  let markStart = from;
  let markEnd = from;
  let foundMark: Mark | undefined;

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child, offset) => {
    const childStart = $pos.start() + offset;
    const childEnd = childStart + child.nodeSize;
    if (from >= childStart && from <= childEnd && child.isText) {
      for (const mark of child.marks) {
        if (mark.type === insertionType || mark.type === deletionType) {
          foundMark = mark;
          markStart = childStart;
          markEnd = childEnd;
        }
      }
    }
  });

  if (foundMark === undefined) {
    const paragraphChange = findParagraphBoundaryChangeAtPosition(state, from);
    return paragraphChange
      ? { from: paragraphChange.from, to: paragraphChange.to }
      : { from, to };
  }

  // Expand to adjacent nodes carrying the *same* mark instance (matching
  // attrs, including revisionId). Two passes — one left-to-right and one
  // right-to-left — so the expansion can cross more than one neighbouring
  // text node on either side (forEach doesn't revisit earlier siblings,
  // which a single-pass walk would need to do to extend leftward by more
  // than one step).
  const sameMark = foundMark;
  const children: {
    childStart: number;
    childEnd: number;
    marks: readonly Mark[];
  }[] = [];
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child, offset) => {
    if (!child.isText) {
      return;
    }
    const childStart = $pos.start() + offset;
    children.push({
      childStart,
      childEnd: childStart + child.nodeSize,
      marks: child.marks,
    });
  });
  let extended = true;
  while (extended) {
    extended = false;
    for (const child of children) {
      if (!child.marks.some((m) => m.eq(sameMark))) {
        continue;
      }
      if (child.childEnd === markStart) {
        markStart = child.childStart;
        extended = true;
      }
      if (child.childStart === markEnd) {
        markEnd = child.childEnd;
        extended = true;
      }
    }
  }

  return { from: markStart, to: markEnd };
}

/**
 * Walk outward from `[fromHint, toHint]` and return the full extent of the
 * tracked-change span carrying `mark` (matched via `Mark.eq`, so attrs like
 * `revisionId` distinguish adjacent changes). Used to keep the
 * navigation/scroll helpers honest when a tracked-change span is split
 * across multiple text nodes due to inline formatting (e.g., a bold word
 * inside an insertion).
 */
function expandTrackedChangeRange(
  state: EditorState,
  mark: Mark,
  fromHint: number,
  toHint: number,
): { from: number; to: number } {
  const carriesSameInlineMark = (node: PMNode | null): node is PMNode =>
    node?.isInline === true && node.marks.some((m) => m.eq(mark));

  // Resolve the boundary positions and hop outward through `nodeBefore`
  // / `nodeAfter` while the neighbouring inline node still carries the
  // same mark instance. O(K) in the number of text nodes that make up
  // the span — `nodesBetween`-based fixed-point expansion is O(K²) and
  // re-walks the same subtree on every iteration.
  let from = fromHint;
  let to = toHint;
  const revisionId: unknown = mark.attrs["revisionId"];
  for (;;) {
    let $from = state.doc.resolve(from);
    let nodeBefore = $from.nodeBefore;
    while (carriesSameInlineMark(nodeBefore)) {
      from -= nodeBefore.nodeSize;
      $from = state.doc.resolve(from);
      nodeBefore = $from.nodeBefore;
    }
    let $to = state.doc.resolve(to);
    let nodeAfter = $to.nodeAfter;
    while (carriesSameInlineMark(nodeAfter)) {
      to += nodeAfter.nodeSize;
      $to = state.doc.resolve(to);
      nodeAfter = $to.nodeAfter;
    }
    // A revision that encloses the control it fills spans the control too:
    // resolving it resolves the control.
    const control = $from.parent;
    if (
      control.type.name !== INLINE_CONTENT_CONTROL_NODE_NAME ||
      typeof revisionId !== "number" ||
      !enclosingRevisionIds(control).includes(revisionId) ||
      from !== $from.start() ||
      to !== $from.end()
    ) {
      return { from, to };
    }
    from = $from.before();
    to = $from.after();
  }
}

/**
 * Find the next tracked change after the given position. Returns the full
 * range of the change (including adjacent text nodes that share the same
 * insertion/deletion mark instance), not just the first text node.
 */
export function findNextChange(
  state: EditorState,
  startPos: number,
): ChangeRange | null {
  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  if (!insertionType && !deletionType) {
    return null;
  }

  const replacementRanges = pairedReplacementRanges(state);
  const result = { value: null as ChangeRange | null };

  state.doc.descendants((node, pos) => {
    if (result.value) {
      return false;
    }
    // Widen from `isText` to `isInline` so an image-only insertion / deletion
    // appears in the find-next walk (an atomic image carries the mark itself,
    // not as a text-node sibling). eigenpal #641.
    if (!node.isInline) {
      return;
    }
    if (pos + node.nodeSize <= startPos) {
      return;
    }

    for (const mark of node.marks) {
      if (mark.type === insertionType || mark.type === deletionType) {
        // Return the FULL expanded range, even when `startPos` lands
        // inside the matched span. A toolbar that does
        // `findNextChange(state, selectionEnd)` and then accepts the
        // returned range must see the whole revision — clamping `from`
        // up to `startPos` truncates the earlier portion of the same
        // change and leaves orphaned marks behind after accept.
        const paired = pairedReplacementRangeAt({
          ranges: replacementRanges,
          from: pos,
          to: pos + node.nodeSize,
        });
        const expanded =
          paired ??
          expandTrackedChangeRange(state, mark, pos, pos + node.nodeSize);
        result.value = {
          from: expanded.from,
          to: expanded.to,
          type: mark.type === insertionType ? "insertion" : "deletion",
        };
        return false;
      }
    }
    return undefined;
  });

  // Wrap around (only once)
  if (result.value === null && startPos > 0) {
    return findNextChange(state, 0);
  }

  return result.value;
}

/**
 * Find the previous tracked change before the given position. Returns the
 * full range of the change (including adjacent text nodes that share the
 * same insertion/deletion mark instance), not just the last text node.
 */
export function findPreviousChange(
  state: EditorState,
  startPos: number,
): ChangeRange | null {
  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  if (!insertionType && !deletionType) {
    return null;
  }

  const replacementRanges = pairedReplacementRanges(state);
  const result = { value: null as ChangeRange | null };
  // Remember the specific mark instance that produced the kept result, so
  // the walk can skip later text nodes covered by the same expansion
  // without skipping a sibling that carries a *different* tracked-change
  // mark (e.g., an `insertion + deletion` overlay where the same text
  // node belongs to two distinct revisions). Skipping by position alone
  // would miss the nearer overlapping change.
  let resultMark: Mark | null = null;

  state.doc.descendants((node, pos) => {
    // Widen from `isText` to `isInline` so an image-only change appears in
    // the find-previous walk. eigenpal #641.
    if (!node.isInline) {
      return;
    }
    if (pos >= startPos) {
      return false;
    }
    if (
      result.value &&
      resultMark &&
      pos < result.value.to &&
      node.marks.every(
        (m) =>
          (m.type !== insertionType && m.type !== deletionType) ||
          m.eq(resultMark!),
      )
    ) {
      // Already covered by the previous expansion AND no additional
      // tracked-change mark sits on this node — safe to skip.
      return;
    }

    for (const mark of node.marks) {
      if (mark.type === insertionType || mark.type === deletionType) {
        const paired = pairedReplacementRangeAt({
          ranges: replacementRanges,
          from: pos,
          to: pos + node.nodeSize,
        });
        const expanded =
          paired ??
          expandTrackedChangeRange(state, mark, pos, pos + node.nodeSize);
        result.value = {
          from: expanded.from,
          to: expanded.to,
          type: mark.type === insertionType ? "insertion" : "deletion",
        };
        resultMark = mark;
      }
    }
    return undefined;
  });

  // Wrap around (only once — guard prevents infinite recursion)
  if (result.value === null && startPos < state.doc.content.size) {
    return findPreviousChange(state, state.doc.content.size);
  }

  return result.value;
}
