/** Comment mark commands and review queries over an editor projection. */
import { type Mark, type Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";

import {
  PARAGRAPH_MARK_CHANGE_KINDS,
  type ParagraphMarkChange,
  type ParagraphMarkChangeKind,
} from "@stll/docx-core/model";

import { sameStatedParagraphNumbering } from "../../docx/numberingReference";
import {
  expectCommentMarkAttrs,
  expectParagraphAttrs,
  expectRunPropertyChangeMarkAttrs,
  expectTableAttrs,
  expectTableCellAttrs,
  expectTableRowAttrs,
  expectTrackedChangeMarkAttrs,
} from "../attrs";
import { enclosingRevisionIds } from "../contentControlRevisions";
import { INLINE_CONTENT_CONTROL_NODE_NAME } from "../extensions/nodes/SdtExtension";
import { getFolioNodeRevisionCarriers } from "../revisionCarriers";
import type { ParagraphPropertyChangeAttrs } from "../schema/nodes";
import { getTableCellMergeChange } from "../tableCellMergeRevision";
import {
  extractTrackedChanges,
  type TrackedChangeEntry,
} from "../utils/extractTrackedChanges";

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

export type TableRowRevisionAttr = {
  revisionId: number;
  author?: string;
  date?: string | null;
};

type TableCellRevisionAttr =
  | {
      kind: "ins" | "del";
      info: {
        revisionId: number;
      };
    }
  | {
      kind: "merge";
      info: {
        revisionId: number;
      };
      verticalMerge?: "continue" | "rest";
      verticalMergeOriginal?: "continue" | "rest";
    };

export function isTableRowRevisionAttr(
  value: unknown,
): value is TableRowRevisionAttr {
  return (
    typeof value === "object" &&
    value !== null &&
    "revisionId" in value &&
    typeof value.revisionId === "number"
  );
}

export function isTableCellRevisionAttr(
  value: unknown,
): value is TableCellRevisionAttr {
  if (
    typeof value !== "object" ||
    value === null ||
    !("kind" in value) ||
    !("info" in value)
  ) {
    return false;
  }
  if (value.kind !== "ins" && value.kind !== "del" && value.kind !== "merge") {
    return false;
  }
  const info = value.info;
  return (
    typeof info === "object" &&
    info !== null &&
    "revisionId" in info &&
    typeof info.revisionId === "number"
  );
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

export function isPPrMarkAttr(value: unknown): value is ParagraphMarkChange {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const kind = (value as { kind?: unknown }).kind;
  const info = (value as { info?: unknown }).info;
  if (!PARAGRAPH_MARK_CHANGE_KINDS.some((allowed) => allowed === kind)) {
    return false;
  }
  if (typeof info !== "object" || info === null) {
    return false;
  }
  return typeof (info as { id?: unknown }).id === "number";
}

/**
 * Whether the mark says the paragraph break was ADDED. A relocation's
 * destination break was added exactly as an insertion's was, and its source
 * break went exactly as a deletion's did; the kinds differ so a reader is told
 * the two ends belong together, not because they resolve differently.
 */
export const paragraphMarkWasAdded = (kind: ParagraphMarkChangeKind): boolean =>
  kind === "ins" || kind === "moveTo";

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

/**
 * Find the document range covered by all revision carriers with any of the
 * given ids. Returns null when none are present (already accepted/rejected, or
 * never existed). A replace operation typically passes two ids; standalone
 * text, formatting, paragraph, section, and table revisions pass a single id.
 */
export function findAIEditRevisionRange(
  state: EditorState,
  revisionIds: number | readonly number[],
): { from: number; to: number } | null {
  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  const runPropertyChangeType = state.schema.marks["runPropertyChange"];
  const idSet = new Set<number>(
    typeof revisionIds === "number" ? [revisionIds] : revisionIds,
  );

  const range = { from: null as number | null, to: null as number | null };
  const includeRange = (from: number, to: number): void => {
    if (range.from === null || from < range.from) {
      range.from = from;
    }
    if (range.to === null || to > range.to) {
      range.to = to;
    }
  };

  state.doc.descendants((node, pos) => {
    for (const carrier of getFolioNodeRevisionCarriers(node, pos)) {
      if (idSet.has(carrier.id)) {
        includeRange(carrier.from, carrier.to);
      }
    }
    // A control an id encloses goes with it, so the range spans the control.
    if (
      node.type.name === INLINE_CONTENT_CONTROL_NODE_NAME &&
      enclosingRevisionIds(node).some((id) => idSet.has(id))
    ) {
      includeRange(pos, pos + node.nodeSize);
    }
    if (node.type.name === "tableRow") {
      for (const attrName of ["trIns", "trDel"] as const) {
        const marker = node.attrs[attrName];
        if (isTableRowRevisionAttr(marker) && idSet.has(marker.revisionId)) {
          includeRange(pos, pos + node.nodeSize);
          return false;
        }
      }
    }
    if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
      const marker = node.attrs["cellMarker"];
      const hasDirectRevision =
        isTableCellRevisionAttr(marker) && idSet.has(marker.info.revisionId);
      const continuationCells = node.attrs["_docxVMergeContinuationCells"];
      const hasCollapsedRevision =
        Array.isArray(continuationCells) &&
        continuationCells.some((cell) => {
          const change = getTableCellMergeChange(cell);
          return change !== null && idSet.has(change.info.id);
        });
      if (hasDirectRevision || hasCollapsedRevision) {
        includeRange(pos, pos + node.nodeSize);
        return false;
      }
    }
    // Widen from `isText` to `isInline` so an AI-edit revision on an inline
    // atom (image, shape) shows up in the matched range. eigenpal #641.
    if (!node.isInline) {
      return;
    }
    for (const mark of node.marks) {
      if (
        mark.type === runPropertyChangeType &&
        expectRunPropertyChangeMarkAttrs(mark).changes.some((change) =>
          idSet.has(change.info.id),
        )
      ) {
        includeRange(pos, pos + node.nodeSize);
        break;
      }
      if (
        (mark.type === insertionType || mark.type === deletionType) &&
        (idSet.has(expectTrackedChangeMarkAttrs(mark).revisionId) ||
          (
            expectTrackedChangeMarkAttrs(mark)._docxRevisionAncestors ?? []
          ).some((ancestor) => idSet.has(ancestor.revisionId)))
      ) {
        includeRange(pos, pos + node.nodeSize);
        break;
      }
    }
    return undefined;
  });

  if (range.from === null || range.to === null) {
    return null;
  }
  return { from: range.from, to: range.to };
}

/**
 * Suggestion (AI-proposed tracked change) commands.
 *
 * A suggestion is a set of `insertion` / `deletion` / `runPropertyChange` marks
 * sharing one `suggestionId` and carrying `provenance: "suggested"`. Accepting
 * rewrites those marks to normal (`"user"`) tracked changes authored by the
 * accepting user; rejecting inverse-applies them like `rejectChange`, reusing
 * the same {@link resolveChange} machinery scoped to the suggestion's revision
 * ids. All three operate as ordinary PM transactions, so they sync via collab
 * and are undoable.
 */

const SUGGESTION_MARK_NAMES = new Set([
  "insertion",
  "deletion",
  "runPropertyChange",
]);

export const isSuggestionMark = (mark: Mark): boolean =>
  mark.attrs["provenance"] === "suggested" &&
  SUGGESTION_MARK_NAMES.has(mark.type.name);

/**
 * The suggestion a comment mark belongs to: a comment a suggestion proposed
 * is pending with it, never saved, accepted and rejected with it.
 */
export const suggestedCommentOf = (mark: Mark): string | null => {
  if (mark.type.name !== "comment") return null;
  const suggestionId = mark.attrs["suggestionId"];
  return typeof suggestionId === "string" && suggestionId.length > 0
    ? suggestionId
    : null;
};

/** The ids of the comments a pending suggestion proposes. */
export function suggestionCommentIds(
  state: EditorState,
  suggestionId: string,
): number[] {
  return [...(collectSuggestions(state).get(suggestionId)?.commentIds ?? [])];
}

export type SuggestionKind =
  | "insertion"
  | "deletion"
  | "formatting"
  | "insertBlock"
  | "insertTable"
  | "insertRow"
  | "deleteRow"
  | "insertColumn"
  | "deleteColumn"
  | "comment";

/**
 * How a suggestion is applied when accepted:
 * - `"tracked"` — converts to a normal OOXML tracked change (`w:ins`/`w:del`,
 *   paragraph-mark `w:ins`, row/cell `w:ins`/`w:del`);
 * - `"direct"` — no OOXML tracked representation exists (a whole inserted
 *   table), so accepting applies it directly; a suggestion that only
 *   proposes a comment is accepted as an ordinary comment;
 * - `"mixed"` — a heterogeneous group (one `suggestionId` spanning a whole
 *   inserted table AND other edits) whose parts apply both ways. Accept still
 *   resolves the whole group in one transaction; this flag lets the host
 *   message that some parts landed directly.
 */
export type SuggestionAppliedAs = "tracked" | "direct" | "mixed";

export type FolioSuggestion = {
  suggestionId: string;
  /** Contiguous document ranges the suggestion covers, in document order. */
  ranges: readonly { from: number; to: number }[];
  /** Which kinds of change the suggestion contains (deduped, in a stable order). */
  kinds: readonly SuggestionKind[];
  /** How accepting this suggestion applies it (see {@link SuggestionAppliedAs}). */
  appliedAs: SuggestionAppliedAs;
};

const suggestionKindOf = (mark: Mark): SuggestionKind => {
  if (mark.type.name === "insertion") {
    return "insertion";
  }
  if (mark.type.name === "deletion") {
    return "deletion";
  }
  return "formatting";
};

/**
 * A node-attr (block/table) suggestion read off a single node: a whole-node
 * `_suggestedInsert`, or a suggested `trIns`/`trDel`/`cellMarker`.
 */
type StructuralSuggestion = {
  suggestionId: string;
  revisionId: number;
  kind: SuggestionKind;
  /** True for a whole-node insert (paragraph/table): rejected by node deletion. */
  isNodeInsert: boolean;
};

export const readSuggestedMarker = (
  marker: unknown,
): { suggestionId: string; revisionId: number } | null => {
  if (
    typeof marker !== "object" ||
    marker === null ||
    (marker as { provenance?: unknown }).provenance !== "suggested"
  ) {
    return null;
  }
  const suggestionId = (marker as { suggestionId?: unknown }).suggestionId;
  const revisionId = (marker as { revisionId?: unknown }).revisionId;
  if (typeof suggestionId !== "string" || typeof revisionId !== "number") {
    return null;
  }
  return { suggestionId, revisionId };
};

export const readSuggestedParagraphPropertyChange = (
  change: ParagraphPropertyChangeAttrs,
): { suggestionId: string; revisionId: number } | null => {
  if (
    change.info.provenance !== "suggested" ||
    typeof change.info.suggestionId !== "string"
  ) {
    return null;
  }
  return { suggestionId: change.info.suggestionId, revisionId: change.info.id };
};

/**
 * The property changes a table, row or cell carries as the model holds them:
 * a pending suggestion's names its suggestion in its info.
 */
const tablePropertyChangesOf = (
  node: PMNode,
): readonly { info: { id: number; suggestionId?: string } }[] => {
  switch (node.type.name) {
    case "table":
      return expectTableAttrs(node).tblPrChange ?? [];
    case "tableRow":
      return expectTableRowAttrs(node).trPrChange ?? [];
    case "tableCell":
    case "tableHeader":
      return expectTableCellAttrs(node).tcPrChange ?? [];
    default:
      return [];
  }
};

type SuggestedParagraphMark = {
  suggestionId: string;
  revisionId: number;
  kind: SuggestionKind;
};

/** The suggestion one paragraph-mark record belongs to, read off its info. */
export const suggestionIdOfMarkRecord = (
  info: ParagraphMarkChange["info"],
): string | null =>
  "provenance" in info &&
  info.provenance === "suggested" &&
  "suggestionId" in info &&
  typeof info.suggestionId === "string"
    ? info.suggestionId
    : null;

const suggestedRecordOf = (
  info: ParagraphMarkChange["info"],
  kind: SuggestionKind,
): SuggestedParagraphMark | null => {
  const suggestionId = suggestionIdOfMarkRecord(info);
  return suggestionId === null
    ? null
    : { suggestionId, revisionId: info.id, kind };
};

/** The suggestion the paragraph mark's own record, its break's, belongs to. */
export const readSuggestedParagraphMark = (
  node: PMNode,
): SuggestedParagraphMark | null => {
  if (node.type.name !== "paragraph") return null;
  const mark = node.attrs["pPrMark"];
  if (!isPPrMarkAttr(mark)) return null;
  return suggestedRecordOf(
    mark.info,
    paragraphMarkWasAdded(mark.kind) ? "insertion" : "deletion",
  );
};

/**
 * Every suggested record of a paragraph's mark: its own, and the later
 * deletion an added mark holds beside its insertion.
 */
const readSuggestedParagraphMarkRecords = (
  node: PMNode,
): SuggestedParagraphMark[] => {
  const own = readSuggestedParagraphMark(node);
  const mark: unknown = node.attrs["pPrMark"];
  const deletion =
    node.type.name === "paragraph" &&
    isPPrMarkAttr(mark) &&
    mark.deletion !== undefined
      ? suggestedRecordOf(mark.deletion.info, "deletion")
      : null;
  return [own, deletion].filter((record) => record !== null);
};

export const readStructuralSuggestion = (
  node: PMNode,
): StructuralSuggestion | null => {
  const attrs = node.attrs;
  // `_suggestedInsert` only carries whole-node semantics for paragraphs
  // (accept → paragraph-mark `w:ins`) and tables (accept → direct). Rows and
  // cells use suggested `trIns`/`trDel`/`cellMarker` instead, so a stray marker
  // on any other node type is ignored rather than mis-classified as a block.
  const name = node.type.name;
  const insertMarker = attrs["_suggestedInsert"];
  if (
    (name === "paragraph" || name === "table") &&
    typeof insertMarker === "object" &&
    insertMarker !== null
  ) {
    const suggestionId = (insertMarker as { suggestionId?: unknown })
      .suggestionId;
    const revisionId = (insertMarker as { revisionId?: unknown }).revisionId;
    if (typeof suggestionId === "string" && typeof revisionId === "number") {
      return {
        suggestionId,
        revisionId,
        kind: name === "table" ? "insertTable" : "insertBlock",
        isNodeInsert: true,
      };
    }
  }
  if (node.type.name === "tableRow") {
    const ins = readSuggestedMarker(attrs["trIns"]);
    if (ins) {
      return { ...ins, kind: "insertRow", isNodeInsert: false };
    }
    const del = readSuggestedMarker(attrs["trDel"]);
    if (del) {
      return { ...del, kind: "deleteRow", isNodeInsert: false };
    }
  }
  if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
    const cellMarker = attrs["cellMarker"];
    if (typeof cellMarker === "object" && cellMarker !== null) {
      // Only insertion/deletion cell markers participate in suggestions.
      // Merge markers never carry suggestion provenance (cell merge/split is
      // `unsupportedMode`), so any other kind is ignored rather than
      // mis-classified as a column insertion.
      const markerKind = (cellMarker as { kind?: unknown }).kind;
      if (markerKind === "ins" || markerKind === "del") {
        const marker = readSuggestedMarker(
          (cellMarker as { info?: unknown }).info,
        );
        if (marker) {
          const kind: SuggestionKind =
            markerKind === "del" ? "deleteColumn" : "insertColumn";
          return { ...marker, kind, isNodeInsert: false };
        }
      }
    }
  }
  return null;
};

type SuggestionAccumulator = {
  segments: { from: number; to: number }[];
  kinds: Set<SuggestionKind>;
  revisionIds: Set<number>;
  /** The comments the suggestion proposes. */
  commentIds: Set<number>;
};

export const collectSuggestions = (
  state: EditorState,
): Map<string, SuggestionAccumulator> => {
  const runPropertyChangeType = state.schema.marks["runPropertyChange"];
  const bySuggestion = new Map<string, SuggestionAccumulator>();
  const entryFor = (suggestionId: string): SuggestionAccumulator => {
    const existing = bySuggestion.get(suggestionId);
    if (existing) {
      return existing;
    }
    const created: SuggestionAccumulator = {
      segments: [],
      kinds: new Set<SuggestionKind>(),
      revisionIds: new Set<number>(),
      commentIds: new Set<number>(),
    };
    bySuggestion.set(suggestionId, created);
    return created;
  };

  state.doc.descendants((node, pos) => {
    const structural = readStructuralSuggestion(node);
    if (structural) {
      const entry = entryFor(structural.suggestionId);
      entry.segments.push({ from: pos, to: pos + node.nodeSize });
      entry.kinds.add(structural.kind);
      entry.revisionIds.add(structural.revisionId);
    }
    for (const { info } of tablePropertyChangesOf(node)) {
      if (info.suggestionId === undefined) continue;
      const entry = entryFor(info.suggestionId);
      entry.segments.push({ from: pos, to: pos + node.nodeSize });
      entry.kinds.add("formatting");
      entry.revisionIds.add(info.id);
    }
    if (node.type.name === "paragraph") {
      for (const paragraphMark of readSuggestedParagraphMarkRecords(node)) {
        const entry = entryFor(paragraphMark.suggestionId);
        entry.segments.push({ from: pos, to: pos + node.nodeSize });
        entry.kinds.add(paragraphMark.kind);
        entry.revisionIds.add(paragraphMark.revisionId);
      }
      const propertyChanges = expectParagraphAttrs(node)._propertyChanges;
      if (Array.isArray(propertyChanges)) {
        for (const change of propertyChanges) {
          const suggested = readSuggestedParagraphPropertyChange(change);
          if (!suggested) {
            continue;
          }
          const entry = entryFor(suggested.suggestionId);
          entry.segments.push({ from: pos, to: pos + node.nodeSize });
          entry.kinds.add("formatting");
          entry.revisionIds.add(suggested.revisionId);
        }
      }
    }
    if (!node.isInline) {
      return undefined;
    }
    for (const mark of node.marks) {
      const commentSuggestion = suggestedCommentOf(mark);
      if (commentSuggestion !== null) {
        const entry = entryFor(commentSuggestion);
        entry.segments.push({ from: pos, to: pos + node.nodeSize });
        entry.kinds.add("comment");
        entry.commentIds.add(expectCommentMarkAttrs(mark).commentId);
        continue;
      }
      if (mark.type.name === "insertion" || mark.type.name === "deletion") {
        // Outer revisions a nested one carries belong to their own
        // suggestions, whatever the nested revision belongs to.
        for (const ancestor of expectTrackedChangeMarkAttrs(mark)
          ._docxRevisionAncestors ?? []) {
          if (ancestor.suggestionId === undefined) continue;
          const entry = entryFor(ancestor.suggestionId);
          entry.segments.push({ from: pos, to: pos + node.nodeSize });
          entry.kinds.add(
            ancestor.type === "insertion" || ancestor.type === "moveTo"
              ? "insertion"
              : "deletion",
          );
          entry.revisionIds.add(ancestor.revisionId);
        }
      }
      if (!isSuggestionMark(mark)) {
        continue;
      }
      const suggestionId = mark.attrs["suggestionId"];
      if (typeof suggestionId !== "string" || suggestionId.length === 0) {
        continue;
      }
      const entry = entryFor(suggestionId);
      entry.segments.push({ from: pos, to: pos + node.nodeSize });
      entry.kinds.add(suggestionKindOf(mark));
      if (mark.type === runPropertyChangeType) {
        for (const change of expectRunPropertyChangeMarkAttrs(mark).changes) {
          entry.revisionIds.add(change.info.id);
        }
      } else if (typeof mark.attrs["revisionId"] === "number") {
        entry.revisionIds.add(mark.attrs["revisionId"]);
      }
    }
    return undefined;
  });

  return bySuggestion;
};

/**
 * A whole inserted table (`insertTable`) has no OOXML tracked representation and
 * accepts directly; every other kind accepts as a tracked change. A group that
 * mixes a table insert with any other kind (a caller stamping heterogeneous
 * operations with one `suggestionId`) is `"mixed"`.
 */
const suggestionAppliedAs = (
  kinds: ReadonlySet<SuggestionKind>,
): SuggestionAppliedAs => {
  // A proposed comment becomes an ordinary comment either way; it decides
  // nothing unless it is all the suggestion holds.
  const changes = [...kinds].filter((kind) => kind !== "comment");
  if (changes.length === 0) {
    return "direct";
  }
  if (!kinds.has("insertTable")) {
    return "tracked";
  }
  const hasOtherKind = changes.some((kind) => kind !== "insertTable");
  return hasOtherKind ? "mixed" : "direct";
};

/** Merge sorted, possibly adjacent/overlapping segments into contiguous ranges. */
export const mergeSegments = (
  segments: readonly { from: number; to: number }[],
): { from: number; to: number }[] => {
  const sorted = [...segments].toSorted(
    (a, b) => a.from - b.from || a.to - b.to,
  );
  const merged: { from: number; to: number }[] = [];
  for (const segment of sorted) {
    const last = merged.at(-1);
    if (last && segment.from <= last.to) {
      last.to = Math.max(last.to, segment.to);
      continue;
    }
    merged.push({ ...segment });
  }
  return merged;
};

const SUGGESTION_KIND_ORDER: readonly SuggestionKind[] = [
  "insertion",
  "deletion",
  "formatting",
  "insertBlock",
  "insertTable",
  "insertRow",
  "deleteRow",
  "insertColumn",
  "deleteColumn",
  "comment",
];

/**
 * List every suggestion in the document for host consumption: its id, the
 * contiguous ranges it covers, which kinds of change it contains, and how
 * accepting it applies (tracked vs direct).
 */
export function getSuggestions(state: EditorState): FolioSuggestion[] {
  const bySuggestion = collectSuggestions(state);
  const suggestions: FolioSuggestion[] = [];
  for (const [suggestionId, entry] of bySuggestion) {
    suggestions.push({
      suggestionId,
      ranges: mergeSegments(entry.segments),
      kinds: SUGGESTION_KIND_ORDER.filter((kind) => entry.kinds.has(kind)),
      appliedAs: suggestionAppliedAs(entry.kinds),
    });
  }
  // Document order by first range start keeps output stable for the host.
  return suggestions.toSorted(
    (a, b) => (a.ranges[0]?.from ?? 0) - (b.ranges[0]?.from ?? 0),
  );
}

/** Physical revision IDs of every pending suggestion, in one document walk. */
export const pendingSuggestionRevisionIds = (
  state: EditorState,
): ReadonlySet<number> => {
  const ids = new Set<number>();
  for (const { revisionIds } of collectSuggestions(state).values())
    for (const id of revisionIds) ids.add(id);
  return ids;
};

/** The first owning suggestion for each physical revision, in source order. */
export const suggestionOwnersByRevision = (
  state: EditorState,
): ReadonlyMap<number, string> => {
  const owners = new Map<number, string>();
  for (const [suggestionId, entry] of collectSuggestions(state)) {
    for (const id of entry.revisionIds)
      if (!owners.has(id)) owners.set(id, suggestionId);
  }
  return owners;
};

/** The suggestion a revision belongs to, or `null` for an ordinary tracked change. */
export function suggestionIdOfRevision(
  state: EditorState,
  revisionId: number,
): string | null {
  for (const [suggestionId, entry] of collectSuggestions(state)) {
    if (entry.revisionIds.has(revisionId)) {
      return suggestionId;
    }
  }
  return null;
}

/**
 * The document range covering every mark belonging to `suggestionId`, plus the
 * revision ids those marks carry. Returns null when the suggestion is absent
 * (already accepted/rejected, or never existed).
 */
export function findSuggestionRange(
  state: EditorState,
  suggestionId: string,
): { from: number; to: number; revisionIds: number[] } | null {
  const entry = collectSuggestions(state).get(suggestionId);
  if (!entry || entry.segments.length === 0) {
    return null;
  }
  const ranges = mergeSegments(entry.segments);
  const from = ranges[0]?.from ?? 0;
  const to = ranges.at(-1)?.to ?? from;
  return { from, to, revisionIds: [...entry.revisionIds] };
}

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
