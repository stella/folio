import type { Mark, Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";

import type { ParagraphMarkChange } from "@stll/docx-core/model";

import {
  expectCommentMarkAttrs,
  expectParagraphAttrs,
  expectRunPropertyChangeMarkAttrs,
  expectTableAttrs,
  expectTableCellAttrs,
  expectTableRowAttrs,
  expectTrackedChangeMarkAttrs,
} from "./attrs";
import { isPPrMarkAttr, paragraphMarkWasAdded } from "./revisionCarriers";
import type { ParagraphPropertyChangeAttrs } from "./schema/nodes";

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

/** The disjoint projected ranges a single pending suggestion covers. */
export const getSuggestionRanges = (
  state: EditorState,
  suggestionId: string,
): { from: number; to: number }[] => {
  const entry = collectSuggestions(state).get(suggestionId);
  return entry === undefined ? [] : mergeSegments(entry.segments);
};

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
