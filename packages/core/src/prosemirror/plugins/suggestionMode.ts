/**
 * Suggestion Mode Plugin
 *
 * When active, intercepts all text insertions and deletions,
 * wrapping them in tracked change marks (insertion/deletion)
 * instead of modifying the document directly.
 *
 * - Typed text is marked as insertion (green underline)
 * - Deleted text is NOT removed — it's marked as deletion (red strikethrough)
 * - Text already marked as insertion by the current author is deleted normally
 *   (retracting your own suggestion)
 */

import {
  joinBackward,
  joinForward,
  selectNodeBackward,
  selectNodeForward,
} from "prosemirror-commands";
import { isHistoryTransaction } from "prosemirror-history";
import { undoInputRule } from "prosemirror-inputrules";
import { Slice, type Node as PMNode, type MarkType } from "prosemirror-model";
import {
  AllSelection,
  EditorState,
  Plugin,
  PluginKey,
  Selection,
  TextSelection,
} from "prosemirror-state";
import type { Transaction } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import { Mapping, ReplaceStep, type Step } from "prosemirror-transform";
import type { EditorView } from "prosemirror-view";

import type { TrackedChangeInfo } from "../../types/document";
import { expectParagraphAttrs, expectTrackedChangeMarkAttrs } from "../attrs";
import { clearIndentOnBackspace } from "../commands/clearParagraphIndent";
import { paragraphPropertiesSnapshot } from "../commands/propertyChangeScope";
import { paragraphEndsItsContainer } from "../containerFinalParagraph";
import type { ParagraphPropertyChangeAttrs } from "../schema/nodes";
import { afterNativeCompositionFlush, handleEditorBeforeInput } from "../textInput";
import { splitBlockClearBorders } from "../extensions/features/BaseKeymapExtension";
import { JOINED_RUNS_RESTYLED_META } from "../extensions/features/JoinedRunStyleExtension";
import { expandNoteReferenceDeletionRange } from "../extensions/marks/noteReferenceDeletion";
import { encloseWholeControls } from "../contentControlRevisions";
import { RUN_FORMATTING_MARK_NAMES } from "../runFormattingMarkNames";
import { canCarryTrackedRunMark } from "../trackedRunInlineAtoms";
import {
  carryParagraphProperties,
  paragraphLeftAfter,
  recordReplacedParagraphProperties,
} from "../paragraphPropertyCarry";
import { rebaseParagraphRuns } from "../rebaseParagraphRunFormatting";
import { paragraphRunStyleContext } from "../runStyleFormatting";
import { getDocumentNumbering } from "./documentNumbering";
import { documentStylesKey, getDocumentStyleResolver } from "./documentStyles";
import { cellPasteRange, pasteTableCells } from "../tableCellPaste";
import { mintRevisionId, seedRevisionIdsFromDoc } from "./revisionIds";

export const suggestionModeKey = new PluginKey<SuggestionModeState>("suggestionMode");
export const SUGGESTION_META = "suggestionModeApplied";
export const SUGGESTION_BYPASS_META = "suggestionModeBypass";
/**
 * The text a tracked text-input transaction typed. Rules that react to typing
 * (list autoformat) read it, because this plugin claims text input before any
 * input rule sees it.
 */
export const SUGGESTED_TEXT_INPUT_META = "suggestionModeTextInput";

type SuggestionModeState = {
  active: boolean;
  author: string;
};

type MarkAttrs = {
  revisionId: number;
  author: string;
  date: string;
};

function makeMarkAttrs(pluginState: SuggestionModeState): MarkAttrs {
  return {
    revisionId: mintRevisionId(),
    author: pluginState.author,
    date: new Date().toISOString(),
  };
}

export function makeRevisionInfo(
  state: EditorState,
): { id: number; author: string; date: string } | null {
  const pluginState = suggestionModeKey.getState(state);
  if (!pluginState?.active) {
    return null;
  }
  const attrs = makeMarkAttrs(pluginState);
  return {
    id: attrs.revisionId,
    author: attrs.author,
    date: attrs.date,
  };
}

/**
 * A removed paragraph mark leaves the paragraph after it once accepted. An
 * edit that reads as the paragraph before it (Backspace or Delete at the
 * break, a deletion running from inside one paragraph into another) gives the
 * paragraph left that paragraph's properties now, under the mark's revision:
 * accepting reads as the edit made directly, rejecting restores its own.
 */
function carryIntoParagraphLeft(
  state: EditorState,
  tr: Transaction,
  source: PMNode,
  survivorPos: number,
  info: TrackedChangeInfo,
): void {
  carryParagraphProperties({
    tr,
    position: survivorPos,
    source,
    styleResolver: getDocumentStyleResolver(state),
    numbering: getDocumentNumbering(state),
    revision: { id: info.id, author: info.author ?? "", date: info.date ?? "" },
  });
}

function makeParagraphMarkInfo(pluginState: SuggestionModeState): TrackedChangeInfo {
  const attrs = makeMarkAttrs(pluginState);
  return {
    id: attrs.revisionId,
    author: attrs.author,
    date: attrs.date,
  };
}

/**
 * Find an adjacent mark of the same type by the same author.
 * Reuses its revisionId so consecutive edits group into one change.
 */
function findAdjacentRevision(
  doc: PMNode,
  pos: number,
  markTypeName: string,
  author: string,
): MarkAttrs | null {
  try {
    const $pos = doc.resolve(pos);
    for (const node of [$pos.nodeBefore, $pos.nodeAfter]) {
      if (node?.isText) {
        const mark = node.marks.find(
          (m) => m.type.name === markTypeName && m.attrs["author"] === author,
        );
        if (mark) {
          return mark.attrs as MarkAttrs;
        }
      }
    }
  } catch {
    /* position out of range */
  }
  return null;
}

/**
 * Find an adjacent revision at either edge of a range.
 * This keeps consecutive backspaces grouped even though the cursor moves left.
 */
function findAdjacentRevisionForRange(
  doc: PMNode,
  from: number,
  to: number,
  markTypeName: string,
  author: string,
): MarkAttrs | null {
  return (
    findAdjacentRevision(doc, from, markTypeName, author) ??
    findAdjacentRevision(doc, to, markTypeName, author)
  );
}

/**
 * Walk a selected range and track its text, paragraph breaks, and table rows.
 * Retract the current author's own inserted text.
 * Processes in reverse order to maintain position validity.
 */
function markRangeAsDeleted(
  tr: Transaction,
  doc: PMNode,
  from: number,
  to: number,
  insertionType: MarkType,
  deletionType: MarkType,
  pluginState: SuggestionModeState,
  /** Set for a deletion the user made across paragraphs (Delete, cut, typing over). */
  joinState?: EditorState,
): void {
  const ranges: { from: number; to: number; isOwnInsert: boolean }[] = [];

  doc.nodesBetween(from, to, (node, pos) => {
    if (!canCarryTrackedRunMark(node)) {
      return;
    }
    const start = Math.max(pos, from);
    const end = Math.min(pos + node.nodeSize, to);
    if (start >= end) {
      return;
    }
    const isOwnInsert = node.marks.some(
      (m) => m.type === insertionType && m.attrs["author"] === pluginState.author,
    );
    // Already struck by someone: re-marking it would overwrite their author,
    // date and revision id with ours, losing who proposed the deletion. The
    // single-character path already steps over such a node; the range path
    // has to agree.
    if (node.marks.some((m) => m.type === deletionType)) {
      return;
    }
    ranges.push({ from: start, to: end, isOwnInsert });
  });

  const delAttrs =
    findAdjacentRevisionForRange(doc, from, to, "deletion", pluginState.author) ||
    makeMarkAttrs(pluginState);

  let removedMark = false;
  doc.nodesBetween(from, to, (node, pos) => {
    if (
      node.type.name === "tableRow" &&
      pos >= from &&
      pos + node.nodeSize <= to &&
      node.attrs["trIns"] == null &&
      node.attrs["trDel"] == null
    ) {
      tr.setNodeAttribute(pos, "trDel", delAttrs);
    }
    if (node.type.name !== "paragraph") {
      return;
    }
    const boundary = pos + node.nodeSize;
    if (boundary > from && boundary <= to && node.attrs["pPrMark"] == null) {
      tr.setNodeAttribute(pos, "pPrMark", {
        kind: "del",
        info: { id: delAttrs.revisionId, author: delAttrs.author, date: delAttrs.date },
      });
      removedMark = true;
    }
  });
  // Words of the first paragraph stay in front of the joined text, so the
  // paragraph left reads as the first one. A deletion that takes the first
  // paragraph whole leaves the last one as it was.
  const $first = doc.resolve(from);
  const $last = doc.resolve(to);
  if (
    joinState &&
    removedMark &&
    $first.parent.type.name === "paragraph" &&
    $last.parent.type.name === "paragraph" &&
    $first.parent !== $last.parent &&
    from > $first.start()
  ) {
    const lastPos = doc === tr.doc ? $last.before() : tr.mapping.map($last.before());
    const lastMark = tr.doc.nodeAt(lastPos)?.attrs["pPrMark"] as { kind?: unknown } | null;
    const lastGoes = lastMark?.kind === "del" || lastMark?.kind === "moveFrom";
    const survivorPos = lastGoes
      ? (paragraphLeftAfter({ doc: tr.doc, paragraphPos: lastPos }) ?? lastPos)
      : lastPos;
    carryIntoParagraphLeft(joinState, tr, $first.parent, survivorPos, {
      id: delAttrs.revisionId,
      author: delAttrs.author,
      date: delAttrs.date,
    });
  }

  if (ranges.length === 0) {
    return;
  }

  // A control the range spans whole is deleted with its text; one the range
  // only empties stays. Recorded before the loop below moves any position.
  encloseWholeControls({
    tr,
    from,
    to,
    revisionId: delAttrs.revisionId,
  });

  for (let i = ranges.length - 1; i >= 0; i--) {
    // SAFETY: i >= 0 and i < ranges.length in for loop
    const range = ranges[i]!;
    if (range.isOwnInsert) {
      tr.delete(range.from, range.to);
    } else {
      tr.addMark(range.from, range.to, deletionType.create(delAttrs));
    }
  }
}

type EnclosePastedRunRevisionsOptions = {
  tr: Transaction;
  from: number;
  to: number;
  revision: MarkAttrs;
};

/** Preserve clipboard revisions inside the revision that inserted their content. */
function enclosePastedRunRevisions({
  tr,
  from,
  to,
  revision: insertion,
}: EnclosePastedRunRevisionsOptions): void {
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (!canCarryTrackedRunMark(node)) return;
    const revision = node.marks.find(
      ({ type }) => type.name === "insertion" || type.name === "deletion",
    );
    if (!revision) return;
    const previous = expectTrackedChangeMarkAttrs(revision);
    tr.addMark(
      Math.max(from, pos),
      Math.min(to, pos + node.nodeSize),
      revision.type.create({
        ...revision.attrs,
        _docxRevisionAncestors: [
          { type: "insertion", ...insertion, outerWrapperCount: 0 },
          ...(previous._docxRevisionAncestors ?? []),
        ],
      }),
    );
  });
}

/**
 * Track every inserted run, paragraph break, and table row in `[from, to)`.
 * Content already carrying a revision keeps its attribution.
 */
function markRangeAsInserted(
  tr: Transaction,
  doc: PMNode,
  from: number,
  to: number,
  insertionType: MarkType,
  deletionType: MarkType,
  attrs: MarkAttrs,
): void {
  doc.nodesBetween(from, to, (node, pos) => {
    if (!canCarryTrackedRunMark(node)) {
      return;
    }
    if (node.marks.some((m) => m.type === insertionType || m.type === deletionType)) {
      return;
    }
    const start = Math.max(pos, from);
    const end = Math.min(pos + node.nodeSize, to);
    if (start >= end) {
      return;
    }
    tr.addMark(start, end, insertionType.create(attrs));
  });
  // A block paste can add paragraph boundaries and whole table rows without
  // adding a markable run at those positions. Track the same structural range
  // as its inline content so rejecting the paste restores the original shape.
  // A range that starts at the end of a paragraph's text and ends after a
  // non-paragraph block (a table) placed that block after the paragraph whole:
  // the paragraph's mark is its own. When the range ends with a paragraph, that
  // paragraph took the original mark and the earlier one is the added break.
  const $end = doc.resolve(to);
  const endsBetweenBlocks =
    !$end.parent.isTextblock &&
    $end.nodeBefore !== null &&
    $end.nodeBefore.type.name !== "paragraph";
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name === "paragraph") {
      const boundary = pos + node.nodeSize;
      const keepsOwnMark = endsBetweenBlocks && boundary - 1 === from;
      if (boundary > from && boundary < to && !keepsOwnMark && node.attrs["pPrMark"] == null) {
        tr.setNodeAttribute(pos, "pPrMark", {
          kind: "ins",
          info: { id: attrs.revisionId, author: attrs.author, date: attrs.date },
        });
      }
    }
    if (
      node.type.name === "tableRow" &&
      pos >= from &&
      pos + node.nodeSize <= to &&
      node.attrs["trIns"] == null &&
      node.attrs["trDel"] == null
    ) {
      tr.setNodeAttribute(pos, "trIns", attrs);
    }
  });
  encloseWholeControls({ tr, from, to, revisionId: attrs.revisionId });
}

/**
 * With track changes on, pasting over a non-empty text selection marks the
 * replaced text as a tracked deletion and the pasted slice as a tracked
 * insertion (so they read as one replacement), matching typing over a
 * selection. Select-all (an `AllSelection`) counts as a text selection over the
 * whole document. Returns false for a collapsed cursor or another selection so the
 * default paste + the `appendTransaction` catch-all marks a plain insertion
 * (eigenpal/docx-editor#784).
 */
export function handleSuggestionPaste(
  view: EditorView,
  slice: Slice,
  pluginState: SuggestionModeState,
): boolean {
  const { selection } = view.state;
  const selectsAll = selection instanceof AllSelection;
  if (!(selection instanceof TextSelection || selectsAll)) {
    return false;
  }
  const insertionType = view.state.schema.marks["insertion"];
  const deletionType = view.state.schema.marks["deletion"];
  if (!insertionType || !deletionType) {
    return false;
  }

  if (selection.empty) {
    const { $from } = selection;
    const closedTable =
      slice.openStart === 0 &&
      slice.openEnd === 0 &&
      slice.content.firstChild?.type.name === "table" &&
      slice.content.lastChild?.type.name === "table";
    const prefix = $from.parent.content.cut(0, $from.parentOffset);
    let deletedPrefix = prefix.size > 0;
    prefix.forEach((node) => {
      if (!node.marks.some((mark) => mark.type === deletionType)) deletedPrefix = false;
    });
    if (!closedTable || $from.parent.type.name !== "paragraph" || !deletedPrefix) {
      return false;
    }
    const suffix = $from.parent.content.cut($from.parentOffset);
    let deletedSuffix = true;
    suffix.forEach((node) => {
      if (!node.marks.some((mark) => mark.type === deletionType)) deletedSuffix = false;
    });
    // The caret is at the visible start of this paragraph. Fit the table
    // before it, as direct paste does after deleting the prefix; fitting at
    // the physical caret instead splits off a paragraph of only struck runs.
    const tr = view.state.tr.setMeta(SUGGESTION_META, true).setMeta("paste", true);
    const at = $from.before();
    tr.replaceRange(at, at, slice);
    const revision = makeMarkAttrs(pluginState);
    const end = tr.mapping.map(at, 1);
    enclosePastedRunRevisions({ tr, from: at, to: end, revision });
    markRangeAsInserted(tr, tr.doc, at, end, insertionType, deletionType, revision);
    if (
      deletedSuffix &&
      $from.parent.attrs["pPrMark"] == null &&
      !paragraphEndsItsContainer($from.doc.resolve($from.before()), "paragraph")
    ) {
      tr.setNodeAttribute(tr.mapping.map($from.before()), "pPrMark", {
        kind: "del",
        info: makeParagraphMarkInfo(pluginState),
      });
    }
    tr.setSelection(TextSelection.create(tr.doc, tr.mapping.map(selection.from)));
    view.dispatch(tr.scrollIntoView());
    return true;
  }

  // Fit open block clipboard edges against the same container boundaries as
  // direct select-all replacement. Fitting beside the struck final paragraph
  // instead can create an empty row at an open table edge.
  if (selectsAll && (slice.openStart > 0 || slice.openEnd > 0)) {
    slice = new Slice(view.state.tr.replaceSelection(slice).doc.content, 0, 0);
  }

  // Select-all spans the block boundaries around the content; replace the text
  // span inside them, as a text selection over the whole document would.
  const { from, to } = selectsAll
    ? { from: Selection.atStart(view.state.doc).from, to: Selection.atEnd(view.state.doc).to }
    : selection;
  const tr = view.state.tr;
  tr.setMeta(SUGGESTION_META, true);

  // A slice that starts with a closed block (a table, finished list items)
  // cannot join the text around the caret, so it splits the textblock it lands
  // in. Over text that starts its textblock, place it before that text: the
  // split then leaves no paragraph behind that only the struck text fills. Over
  // whole textblocks, strike their paragraph marks too, so accepting removes
  // them as replacing them directly would.
  const { doc } = view.state;
  const $from = doc.resolve(from);
  const $to = doc.resolve(to);
  const closedBlocks = slice.openStart === 0 && slice.content.firstChild?.isBlock === true;
  const beforeStruck = closedBlocks && $from.parent.isTextblock && $from.parentOffset === 0;
  const wholeBlocks =
    beforeStruck && $to.parent.isTextblock && $to.parentOffset === $to.parent.content.size;
  // The paragraph a container ends with never carries a tracked mark, so it
  // cannot be struck whole. Only its words are struck, and a table goes in
  // front of it: accepting leaves the table and the emptied paragraph after
  // it, which still ends the container. Closing paragraphs take its place
  // instead, see `rotateIntoFinalParagraph`.
  const endsContainer =
    wholeBlocks && paragraphEndsItsContainer(doc.resolve($to.before()), $to.parent.type.name);
  const rotates =
    endsContainer &&
    slice.content.firstChild?.type === $to.parent.type &&
    slice.content.lastChild?.type === $to.parent.type;

  // 1. Strike through the replaced selection (or retract own pending inserts).
  //    A table going in front of a container's last paragraph is placed
  //    first: struck before, the paragraphs it lands in front of would lose
  //    their struck marks' undo to the insertion at their position.
  const tableFirst = endsContainer && !rotates;
  if (!tableFirst) {
    markRangeAsDeleted(
      tr,
      doc,
      from,
      wholeBlocks && !endsContainer ? $to.after() : to,
      insertionType,
      deletionType,
      pluginState,
    );
  }

  // A paste over words running from inside one paragraph into another removes
  // the first one's break: the paragraph its words now run on into takes its
  // properties, under that break's revision, as a deletion across the break
  // does. Pasted words go in after, so they read in those properties as they
  // are; pasted paragraphs first, so the paragraph left is found past them.
  const carryAcross = () => {
    if (
      selectsAll ||
      $from.parent.type.name !== "paragraph" ||
      $from.parent === $to.parent ||
      from <= $from.start()
    ) {
      return;
    }
    const firstPos = tr.mapping.map($from.before());
    const mark = tr.doc.nodeAt(firstPos)?.attrs["pPrMark"] as ParagraphMarkAttr | null | undefined;
    const survivorPos = paragraphLeftAfter({ doc: tr.doc, paragraphPos: firstPos });
    if (mark?.kind === "del" && survivorPos !== null) {
      carryIntoParagraphLeft(view.state, tr, $from.parent, survivorPos, mark.info);
    }
  };
  const inlineSlice = slice.content.firstChild?.isInline === true;
  if (inlineSlice) {
    carryAcross();
  }

  // 2. Insert the pasted slice beside the struck-through selection.
  //    `replaceRange` fits the slice's open sides into the surrounding content
  //    the way ProseMirror's normal paste does, so block clipboard content (a
  //    copied table or whole paragraphs) is placed structurally instead of
  //    failing or dropping nodes as a raw `replace` at an inline point would.
  let at = tr.mapping.map(to);
  if (rotates) {
    at = tr.mapping.map($to.after());
  } else if (tableFirst) {
    at = $from.before();
  } else if (beforeStruck) {
    at = tr.mapping.map(from, -1);
  }
  const firstStep = tr.steps.length;
  tr.replaceRange(at, at, slice);
  // The range the paste occupies is where its steps put content, which for
  // block content is past the textblock boundary rather than at the caret.
  const placed = tr.steps.slice(firstStep).flatMap((step, offset) => {
    const later = new Mapping(tr.mapping.maps.slice(firstStep + offset + 1));
    const ranges: { from: number; to: number }[] = [];
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror StepMap.forEach
    step.getMap().forEach((_oldFrom, _oldTo, newFrom, newTo) => {
      // oxlint-disable-next-line unicorn/no-array-method-this-argument -- ProseMirror Mapping.map(pos, assoc)
      ranges.push({ from: later.map(newFrom, -1), to: later.map(newTo, 1) });
    });
    return ranges;
  });
  let insertFrom = placed.length === 0 ? at : Math.min(...placed.map((range) => range.from));
  let insertTo = placed.length === 0 ? at : Math.max(...placed.map((range) => range.to));
  if (tableFirst) {
    const struck = tr.steps.length;
    markRangeAsDeleted(
      tr,
      tr.doc,
      tr.mapping.map(from),
      tr.mapping.map(to),
      insertionType,
      deletionType,
      pluginState,
    );
    const since = tr.mapping.slice(struck);
    insertFrom = since.map(insertFrom, -1);
    insertTo = since.map(insertTo, -1);
  }

  // 3. Track the paste around any revisions the clipboard already carried.
  const insertAttrs =
    findAdjacentRevision(doc, from, "insertion", pluginState.author) || makeMarkAttrs(pluginState);
  enclosePastedRunRevisions({ tr, from: insertFrom, to: insertTo, revision: insertAttrs });
  markRangeAsInserted(tr, tr.doc, insertFrom, insertTo, insertionType, deletionType, insertAttrs);

  if (!inlineSlice) {
    carryAcross();
  }

  // An open paste splits the paragraph at its end: its original break ends
  // the last pasted part. Rejecting the inserted breaks leaves that part, so
  // it must record the end paragraph's properties even for a cross-paragraph
  // replacement. A carry onto the first pasted part cannot survive that join.
  if (!closedBlocks && $to.parent.type.name === "paragraph") {
    const $end = tr.doc.resolve(tr.mapping.map(to));
    if ($end.parent.type.name === "paragraph" && $end.before() !== tr.mapping.map($from.before())) {
      recordReplacedParagraphProperties({
        tr,
        position: $end.before(),
        replaced: $to.parent,
        revision: {
          id: insertAttrs.revisionId,
          author: insertAttrs.author,
          date: insertAttrs.date,
        },
      });
    }
  }

  // Replacing the whole document with an open slice takes the first pasted
  // paragraph's properties. Inserting beside the struck selection instead
  // initially keeps the old final paragraph's properties on that first part.
  const firstPasted = slice.content.firstChild;
  if (selectsAll && !closedBlocks && firstPasted?.type.name === "paragraph") {
    const $first = tr.doc.resolve(insertFrom);
    if ($first.parent.type.name === "paragraph") {
      carryParagraphProperties({
        tr,
        position: $first.before(),
        source: firstPasted,
        revision: {
          id: insertAttrs.revisionId,
          author: insertAttrs.author,
          date: insertAttrs.date,
        },
        styleResolver: getDocumentStyleResolver(view.state),
        numbering: getDocumentNumbering(view.state),
      });
    }
  }

  if (rotates) {
    const rotation = tr.steps.length;
    const joined = rotateIntoFinalParagraph(
      view.state,
      tr,
      tr.mapping.map($to.before()),
      insertAttrs,
    );
    insertFrom = Math.min(tr.mapping.slice(rotation).map(insertFrom, -1), joined);
    insertTo = tr.mapping.slice(rotation).map(insertTo);
  }

  // Collapse to the end of the pasted content. Without this the struck-through
  // original plus the pasted text stay selected, so the next keystroke would
  // delete them both. `near` keeps the selection valid even when the pasted
  // slice was block content and `insertTo` lands at a block boundary.
  tr.setSelection(TextSelection.near(tr.doc.resolve(insertTo)));

  view.dispatch(tr.scrollIntoView());
  return true;
}

/**
 * With track changes on, a paste into a table's cells (a cell selection, or a
 * block of copied cells) replaces each cell's content the way a paste over a
 * text selection does: the old content struck through, the pasted content
 * after it as an insertion. When the pasted content is whole paragraphs, the
 * last one ends the cell, so its paragraph mark is tracked as added too:
 * rejecting then drops it and accepting joins the struck paragraphs into it.
 */
export function handleSuggestionTableCellPaste(
  view: EditorView,
  slice: Slice,
  pluginState: SuggestionModeState,
): boolean {
  const insertionType = view.state.schema.marks["insertion"];
  const deletionType = view.state.schema.marks["deletion"];
  if (!insertionType || !deletionType) {
    return false;
  }
  const revision = makeMarkAttrs(pluginState);
  return pasteTableCells(
    view.state,
    slice,
    (tr) => view.dispatch(tr.setMeta(SUGGESTION_META, true)),
    {
      revision,
      replaceCellContent: (tr, cellPos, content) => {
        const cell = tr.doc.nodeAt(cellPos);
        if (!cell) {
          return;
        }
        const { from, to } = cellPasteRange(cell, cellPos, content);
        const mapFrom = tr.mapping.maps.length;
        markRangeAsDeleted(tr, tr.doc, from, to, insertionType, deletionType, pluginState);
        const insertFrom = tr.mapping.slice(mapFrom).map(to);
        const sizeBefore = tr.doc.content.size;
        tr.replaceRange(insertFrom, insertFrom, content);
        let insertTo = insertFrom + (tr.doc.content.size - sizeBefore);
        let pastedFrom = insertFrom;
        // An open paste splits the cell's final paragraph. Its old closing
        // mark belongs to the last pasted part; the new break before that is
        // an insertion, even when the old paragraph was already struck.
        const finalParagraph = cell.lastChild;
        const $pasteFrom = tr.doc.resolve(pastedFrom);
        const $pasteTo = tr.doc.resolve(insertTo);
        if (
          content.openStart > 0 &&
          content.openEnd > 0 &&
          finalParagraph?.type.name === "paragraph" &&
          $pasteFrom.parent.type.name === "paragraph" &&
          $pasteTo.parent.type.name === "paragraph" &&
          $pasteFrom.before() !== $pasteTo.before()
        ) {
          tr.setNodeAttribute($pasteFrom.before(), "pPrMark", {
            kind: "ins",
            info: { id: revision.revisionId, author: revision.author, date: revision.date },
          });
          tr.setNodeAttribute($pasteTo.before(), "pPrMark", finalParagraph.attrs["pPrMark"]);
          recordReplacedParagraphProperties({
            tr,
            position: $pasteTo.before(),
            replaced: finalParagraph,
            revision: {
              id: revision.revisionId,
              author: revision.author,
              date: revision.date,
            },
          });
        }
        // Retracting our pasted words also retracts their empty paragraphs.
        // Keep paragraphs carrying original struck content for rejection; the
        // new paste now supplies the required final paragraph of the cell.
        const emptyInsertions: { from: number; to: number }[] = [];
        const replacedCell = tr.doc.nodeAt(cellPos);
        // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
        replacedCell?.forEach((node, offset) => {
          const pos = cellPos + 1 + offset;
          if (
            pos < pastedFrom &&
            node.type.name === "paragraph" &&
            node.content.size === 0 &&
            isCurrentAuthorParagraphInsertion(node.attrs["pPrMark"], pluginState.author)
          ) {
            emptyInsertions.push({ from: pos, to: pos + node.nodeSize });
          }
        });
        const retraction = tr.mapping.maps.length;
        for (const range of emptyInsertions.toReversed()) tr.delete(range.from, range.to);
        const retracted = tr.mapping.slice(retraction);
        pastedFrom = retracted.map(pastedFrom);
        insertTo = retracted.map(insertTo);
        enclosePastedRunRevisions({ tr, from: pastedFrom, to: insertTo, revision });
        markRangeAsInserted(
          tr,
          tr.doc,
          pastedFrom,
          insertTo,
          insertionType,
          deletionType,
          revision,
        );
        tr.doc.nodesBetween(pastedFrom, insertTo, (node, pos) => {
          if (
            node.type.name === "paragraph" &&
            pos >= pastedFrom &&
            pos + node.nodeSize <= insertTo &&
            node.attrs["pPrMark"] == null
          ) {
            tr.setNodeAttribute(pos, "pPrMark", {
              kind: "ins",
              info: { id: revision.revisionId, author: revision.author, date: revision.date },
            });
          }
        });
      },
    },
  );
}

/**
 * Pasted paragraphs replaced a container's last paragraph. They were placed
 * after it, and the paragraph was struck but keeps its mark, as a container's
 * last paragraph must. Join the first pasted paragraph onto it, so the struck
 * text and the first paste share one paragraph whose break is the inserted one,
 * and record the replaced paragraph's formatting on the paragraph that now ends
 * the container: accepting leaves only the paste, and rejecting closes every
 * pasted break back into one paragraph that returns to its old formatting.
 * Returns the position of the joined paragraph.
 */
function rotateIntoFinalParagraph(
  state: EditorState,
  tr: Transaction,
  replacedPos: number,
  attrs: MarkAttrs,
): number {
  const replaced = tr.doc.nodeAt(replacedPos);
  const first = replaced ? tr.doc.nodeAt(replacedPos + replaced.nodeSize) : null;
  if (!replaced || !first) {
    return replacedPos;
  }
  const previousFormatting = paragraphPropertiesSnapshot(replaced);
  const { _sectionProperties: sectionProperties } = replaced.attrs;
  tr.join(replacedPos + replaced.nodeSize);
  tr.setNodeMarkup(replacedPos, undefined, first.attrs);
  // The struck words now sit in the first pasted paragraph: what the replaced
  // paragraph's style lent them is re-read in that one's.
  const styleResolver = getDocumentStyleResolver(state);
  if (styleResolver && replaced.content.size > 0) {
    rebaseParagraphRuns({
      previousContext: paragraphRunStyleContext(replaced, styleResolver),
      paragraphPosition: replacedPos,
      range: { from: 0, to: replaced.content.size },
      styleResolver,
      tr,
    });
  }

  // The paragraph that now ends the container carries the replaced one's
  // section, and no mark. Rejecting the pasted breaks closes every pasted
  // paragraph into it, the one whose break stays, so the change recorded here
  // returns it to the replaced paragraph's formatting. It is the paragraph's
  // only one: a change the pasted paragraph brought records formatting the
  // document never had here, and a paragraph holds one w:pPrChange.
  const $joined = tr.doc.resolve(replacedPos);
  const container = $joined.parent;
  let finalPos = replacedPos;
  for (let index = $joined.index(); index < container.childCount - 1; index += 1) {
    finalPos += container.child(index).nodeSize;
  }
  const final = tr.doc.nodeAt(finalPos);
  if (final) {
    tr.setNodeMarkup(finalPos, undefined, {
      ...final.attrs,
      pPrMark: null,
      ...(sectionProperties == null ? {} : { _sectionProperties: sectionProperties }),
      _propertyChanges: [
        {
          type: "paragraphPropertyChange",
          info: { id: attrs.revisionId, author: attrs.author, date: attrs.date },
          previousFormatting,
        } satisfies ParagraphPropertyChangeAttrs,
      ],
    });
  }
  return replacedPos;
}

/**
 * Insert text as a tracked insertion, optionally marking replaced selection as deletion.
 */
function applySuggestionInsert(
  view: EditorView,
  from: number,
  to: number,
  text: string,
  pluginState: SuggestionModeState,
): boolean {
  const insertionType = view.state.schema.marks["insertion"];
  if (!insertionType) {
    return false;
  }

  const tr = view.state.tr;
  tr.setMeta(SUGGESTION_META, true);
  tr.setMeta(SUGGESTED_TEXT_INPUT_META, text);

  // Select-all spans the block boundaries around the content; typing replaces
  // the text span inside them, as over a text selection of the whole document:
  // the typed text lands in the last paragraph, whose break is the one that
  // stays, formatted as the first character it replaces.
  const selectsAll = view.state.selection instanceof AllSelection;
  if (selectsAll) {
    from = Math.max(from, Selection.atStart(view.state.doc).from);
    to = Math.min(to, Selection.atEnd(view.state.doc).to);
  }
  const replacedFormatting = selectsAll
    ? (view.state.doc.resolve(from).nodeAfter?.marks ?? []).filter(({ type }) =>
        RUN_FORMATTING_MARK_NAMES.has(type.name),
      )
    : null;

  const insertAttrs =
    findAdjacentRevision(view.state.doc, from, "insertion", pluginState.author) ||
    makeMarkAttrs(pluginState);

  if (from !== to) {
    const deletionType = view.state.schema.marks["deletion"];
    if (deletionType) {
      markRangeAsDeleted(
        tr,
        view.state.doc,
        from,
        to,
        insertionType,
        deletionType,
        pluginState,
        view.state,
      );
    }
  }

  const insertAt = tr.mapping.map(to);
  tr.insertText(text, insertAt, insertAt);

  // Strip inherited deletion marks — new text must never be marked as deleted.
  const deletionType = view.state.schema.marks["deletion"];
  if (deletionType) {
    tr.removeMark(insertAt, insertAt + text.length, deletionType);
  }

  // Apply the correct insertion mark. If the cursor was inside an existing
  // insertion by the same author, insertText already inherited that mark and
  // insertAttrs will match — addMark is effectively a no-op that preserves
  // the continuous mark span. We intentionally do NOT removeMark(insertionType)
  // first, because that fragments the mark span and creates a nested change.
  tr.addMark(insertAt, insertAt + text.length, insertionType.create(insertAttrs));
  if (replacedFormatting) {
    for (const mark of tr.doc.resolve(insertAt).nodeAfter?.marks ?? []) {
      if (RUN_FORMATTING_MARK_NAMES.has(mark.type.name)) {
        tr.removeMark(insertAt, insertAt + text.length, mark.type);
      }
    }
    for (const mark of replacedFormatting) {
      tr.addMark(insertAt, insertAt + text.length, mark);
    }
  }
  separateSplitStretch(tr, insertAt, insertAt + text.length, insertAttrs.revisionId);
  if (view.state.selection instanceof AllSelection) {
    // The caret follows the typed text, as after typing over any selection.
    tr.setSelection(TextSelection.create(tr.doc, insertAt + text.length));
  }

  view.dispatch(tr.scrollIntoView());
  return true;
}

/**
 * Text typed inside another revision leaves that revision in two stretches.
 * A saved package writes each stretch as a wrapper of its own with its own
 * `w:id`, so the stretch after the typed text gets a fresh id now: the editor
 * lists, resolves and saves the same changes a reopened package does.
 */
function separateSplitStretch(tr: Transaction, from: number, to: number, typedId: number): void {
  const $from = tr.doc.resolve(from);
  const $to = tr.doc.resolve(to);
  const before = $from.nodeBefore;
  if (!before || $to.parent !== $from.parent) {
    return;
  }
  const parent = $to.parent;
  const start = $to.start();
  for (const mark of before.marks) {
    if (mark.type.name !== "insertion" && mark.type.name !== "deletion") {
      continue;
    }
    const revisionId = mark.attrs["revisionId"];
    if (revisionId === typedId || typeof revisionId !== "number") {
      continue;
    }
    // The stretch after the typed text: the same revision, uninterrupted.
    let end = to;
    let offset = to - start;
    while (offset < parent.content.size) {
      const child = parent.childAfter(offset).node;
      if (!child?.marks.some((candidate) => candidate.eq(mark))) {
        break;
      }
      offset += child.nodeSize;
      end = start + offset;
    }
    if (end === to) {
      continue;
    }
    tr.removeMark(to, end, mark);
    tr.addMark(to, end, mark.type.create({ ...mark.attrs, revisionId: mintRevisionId() }));
  }
}

/**
 * Remove `from`..`to` of `tr.doc` as a suggestion: text the current author
 * inserted is retracted, anything else is marked deleted. `false` when the
 * editor is not suggesting.
 */
export function suggestRangeDeletion(
  state: EditorState,
  tr: Transaction,
  from: number,
  to: number,
): boolean {
  const pluginState = suggestionModeKey.getState(state);
  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  if (!pluginState?.active || !insertionType || !deletionType) {
    return false;
  }
  markRangeAsDeleted(tr, tr.doc, from, to, insertionType, deletionType, pluginState);
  return true;
}

/**
 * Track-changes Enter: split the paragraph and stamp `pPrMark = 'ins'` on the
 * FIRST half of the split (the upper paragraph), per ECMA-376 §17.13.5. The
 * actual splitting reuses `splitBlockClearBorders` so border/style/stored-mark
 * behavior stays identical to ordinary editing.
 *
 * If the source paragraph already carries a `pPrMark`, leave it alone — a
 * prior author's revision must not be silently overwritten.
 */
export function handleSuggestionEnter(view: EditorView, pluginState: SuggestionModeState): boolean {
  const state = view.state;
  if (state.selection.$from.parent.type.name !== "paragraph") {
    return false;
  }
  const replacement = { tr: null as Transaction | null };
  let splitState = state;
  if (state.selection instanceof TextSelection && !state.selection.empty) {
    // Enter replaces selected text, but rejecting must retain it. Prepare the
    // deletion without dispatching so replacement and split form one undo event.
    const deleted = handleSuggestionDelete(
      state,
      (tr) => {
        replacement.tr = tr;
        // The split reads only the style resolver from plugin state. Preserve
        // its explicit key while omitting transaction hooks from the preview.
        const styles = documentStylesKey.get(state);
        const plugins = [];
        if (styles) {
          const spec = { ...styles.spec };
          delete spec.appendTransaction;
          delete spec.filterTransaction;
          plugins.push(new Plugin(spec));
        }
        const preview = state.reconfigure({ plugins });
        splitState = preview.apply(tr);
      },
      "forward",
    );
    if (!deleted) return false;
  }
  const { $from } = splitState.selection;
  const sourcePos = $from.before();
  const sourceAttrs = $from.parent.attrs;

  const captured = { tr: null as Transaction | null };
  const ok = splitBlockClearBorders(
    splitState,
    (tr: Transaction) => {
      captured.tr = tr;
    },
    view,
  );
  if (!ok || !captured.tr) {
    return false;
  }
  const tr = replacement.tr ?? captured.tr;
  if (replacement.tr) {
    for (const step of captured.tr.steps) tr.step(step);
    tr.setSelection(captured.tr.selection.getBookmark().resolve(tr.doc));
    tr.setStoredMarks(captured.tr.storedMarks);
  }
  tr.setMeta(SUGGESTION_META, true);
  if (sourceAttrs["pPrMark"] == null) {
    const sourceParagraph = tr.doc.nodeAt(sourcePos);
    if (sourceParagraph?.type.name === "paragraph") {
      const markInfo: ParagraphMarkAttr = {
        kind: "ins",
        info: makeParagraphMarkInfo(pluginState),
      };
      tr.setNodeAttribute(sourcePos, "pPrMark", markInfo);
      // The new paragraph holds the source's break: rejecting the inserted
      // one leaves it, so it records what the source read as.
      const $caret = tr.selection.$from;
      if ($caret.parent.type.name === "paragraph" && $caret.before() !== sourcePos) {
        recordReplacedParagraphProperties({
          tr,
          position: $caret.before(),
          replaced: $from.parent,
          revision: markInfo.info,
        });
      }
    }
  }
  view.dispatch(tr.scrollIntoView());
  return true;
}

type ParagraphMarkAttr = {
  kind: "ins" | "del";
  info: TrackedChangeInfo;
};

type CaretDeleteTarget =
  | { type: "inlineUnit"; from: number; to: number }
  | { type: "paragraphEdge" };

/** Find the adjacent visible unit without consuming a bookmark range marker. */
function caretDeleteTarget(
  state: EditorState,
  direction: "backward" | "forward",
): CaretDeleteTarget {
  const { $from } = state.selection;
  const backward = direction === "backward";
  let from = backward ? $from.pos - 1 : $from.pos;
  let to = backward ? $from.pos : $from.pos + 1;
  while (from >= $from.start() && to <= $from.end()) {
    const adjacent = state.doc.resolve(from).nodeAfter;
    if (adjacent?.type.name !== "bookmarkBoundary") {
      return { type: "inlineUnit", from, to };
    }
    from += backward ? -adjacent.nodeSize : adjacent.nodeSize;
    to += backward ? -adjacent.nodeSize : adjacent.nodeSize;
  }
  return { type: "paragraphEdge" };
}

/**
 * Detect a caret-at-paragraph-boundary scenario where Backspace/Delete should
 * record a tracked paragraph-mark deletion instead of joining paragraphs.
 *
 * Returns the position of the paragraph whose `pPrMark` should be set, or
 * `null` if the caret is mid-paragraph (let the normal text-delete path run).
 *
 * - Backspace-at-paragraph-start → previous sibling paragraph.
 * - Delete-at-paragraph-end     → current paragraph.
 *
 * Returns `null` when there is no adjacent sibling paragraph (the join would
 * not produce paragraph-mark merging, e.g. doc start, doc end, or a table).
 * Bookmark markers at an edge do not hide the adjacent paragraph break.
 */
export function paragraphBoundaryTarget(
  state: EditorState,
  direction: "backward" | "forward",
): number | null {
  const { $from, empty } = state.selection;
  if (!empty) {
    return null;
  }
  if ($from.parent.type.name !== "paragraph") {
    return null;
  }
  const paragraphStart = $from.before();
  const paragraphEnd = $from.after();

  if (caretDeleteTarget(state, direction).type !== "paragraphEdge") {
    return null;
  }

  if (direction === "backward") {
    if (paragraphStart === 0) {
      return null;
    }
    const $prevEdge = state.doc.resolve(paragraphStart);
    const prev = $prevEdge.nodeBefore;
    if (!prev || prev.type.name !== "paragraph") {
      return null;
    }
    return paragraphStart - prev.nodeSize;
  }

  if (paragraphEnd >= state.doc.content.size) {
    return null;
  }
  const $nextEdge = state.doc.resolve(paragraphEnd);
  const next = $nextEdge.nodeAfter;
  if (!next || next.type.name !== "paragraph") {
    return null;
  }
  return paragraphStart;
}

/**
 * Mark a paragraph break as a tracked deletion. `target` is the position of
 * the paragraph whose closing mark is being deleted (per OOXML, that's the
 * previous paragraph for Backspace-at-start and the current paragraph for
 * Delete-at-end). If the existing mark is the current author's insertion,
 * retract it by joining the paragraphs back together.
 */
function applyPPrDel(
  view: EditorView,
  targetParagraphPos: number,
  pluginState: SuggestionModeState,
): boolean {
  const targetNode = view.state.doc.nodeAt(targetParagraphPos);
  if (!targetNode || targetNode.type.name !== "paragraph") {
    return false;
  }
  const existingMark = targetNode.attrs["pPrMark"];
  if (isCurrentAuthorParagraphInsertion(existingMark, pluginState.author)) {
    const tr = view.state.tr;
    tr.setMeta(SUGGESTION_META, true);
    const joinPos = targetParagraphPos + targetNode.nodeSize;
    const joined = view.state.doc.nodeAt(joinPos);
    try {
      tr.join(joinPos);
      // Retracting this break leaves the following paragraph's closing mark.
      // Keep the editing join's formatting, but retain what that paragraph
      // read as so rejecting the remaining inserted breaks restores it.
      tr.setNodeAttribute(
        targetParagraphPos,
        "pPrMark",
        joined?.type === targetNode.type ? expectParagraphAttrs(joined).pPrMark : null,
      );
      if (joined?.type === targetNode.type) {
        recordReplacedParagraphProperties({
          tr,
          position: targetParagraphPos,
          replaced: joined,
          revision: existingMark.info,
        });
      }
      // The paragraph keeps its own properties, so the words the retraction
      // brings in drop what their old paragraph's style lent them and read in
      // this one's: a run with no formatting of its own stays without any.
      const styleResolver = getDocumentStyleResolver(view.state);
      if (styleResolver && joined?.type === targetNode.type && joined.content.size > 0) {
        rebaseParagraphRuns({
          previousContext: paragraphRunStyleContext(joined, styleResolver),
          paragraphPosition: targetParagraphPos,
          range: {
            from: targetNode.content.size,
            to: targetNode.content.size + joined.content.size,
          },
          styleResolver,
          tr,
        });
      }
      tr.setMeta(JOINED_RUNS_RESTYLED_META, true);
      view.dispatch(tr.scrollIntoView());
    } catch {
      return true;
    }
    return true;
  }
  if (existingMark != null) {
    return true;
  }
  const tr = view.state.tr;
  tr.setMeta(SUGGESTION_META, true);
  const markInfo: ParagraphMarkAttr = {
    kind: "del",
    info: makeParagraphMarkInfo(pluginState),
  };
  tr.setNodeAttribute(targetParagraphPos, "pPrMark", markInfo);
  carryIntoParagraphLeft(
    view.state,
    tr,
    targetNode,
    paragraphLeftAfter({ doc: tr.doc, paragraphPos: targetParagraphPos }) ??
      targetParagraphPos + targetNode.nodeSize,
    markInfo.info,
  );
  view.dispatch(tr.scrollIntoView());
  return true;
}

function isCurrentAuthorParagraphInsertion(
  value: unknown,
  author: string,
): value is ParagraphMarkAttr {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const mark = value as { kind?: unknown; info?: unknown };
  if (mark.kind !== "ins" || typeof mark.info !== "object" || mark.info === null) {
    return false;
  }
  const info = mark.info as { author?: unknown };
  return info.author === author;
}

/**
 * Handle delete (forward or backward) in suggestion mode.
 */
function handleSuggestionDelete(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  direction: "backward" | "forward",
): boolean {
  const pluginState = suggestionModeKey.getState(state);
  if (!pluginState?.active) {
    return false;
  }

  const { $from, $to, empty } = state.selection;
  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  if (!insertionType || !deletionType) {
    return false;
  }

  if (!dispatch) {
    return true;
  }

  const tr = state.tr;
  tr.setMeta(SUGGESTION_META, true);

  // --- Selection delete ---
  if (!empty) {
    const ranges = state.selection.ranges
      .map(
        ({ $from: rangeFrom, $to: rangeTo }) =>
          expandNoteReferenceDeletionRange(state.doc, rangeFrom.pos, rangeTo.pos) ?? {
            from: rangeFrom.pos,
            to: rangeTo.pos,
          },
      )
      .sort((a, b) => b.from - a.from);
    for (const { from, to } of ranges) {
      markRangeAsDeleted(tr, state.doc, from, to, insertionType, deletionType, pluginState, state);
    }
    // Clearing cells preserves the rectangle, as the table deletion command
    // does, so the next input still acts on those cells. Text deletion instead
    // collapses the cursor after the marked/retracted content.
    if (state.selection instanceof CellSelection) {
      tr.setSelection(state.selection.map(tr.doc, tr.mapping));
    } else {
      const cursorPos = tr.mapping.map($to.pos);
      tr.setSelection(TextSelection.near(tr.doc.resolve(cursorPos)));
    }
    dispatch(tr.scrollIntoView());
    return true;
  }

  // --- Caret delete (one character or a whole note reference) ---
  const isBackward = direction === "backward";
  const target = caretDeleteTarget(state, direction);
  if (target.type === "paragraphEdge") {
    const edge = isBackward ? $from.start() : $from.end();
    if ($from.pos === edge) return false;
    // The paragraph-break handler claimed paragraph joins above. For another
    // block sibling, run the ordinary join at the visible edge. A clean
    // transient state avoids dispatching a selection-only transaction first.
    const virtual = EditorState.create({
      doc: state.doc,
      selection: TextSelection.create(state.doc, edge),
    });
    const join = isBackward ? joinBackward : joinForward;
    if (join(virtual, dispatch)) return true;
    const selectNode = isBackward ? selectNodeBackward : selectNodeForward;
    return selectNode(virtual, dispatch);
  }
  const deletePos = target.from;
  const deleteEnd = target.to;

  const noteRange = expandNoteReferenceDeletionRange(state.doc, deletePos, deleteEnd);
  const rangeFrom = noteRange?.from ?? deletePos;
  const rangeTo = noteRange?.to ?? deleteEnd;
  const $deletePos = state.doc.resolve(rangeFrom);
  const nodeAfter = $deletePos.nodeAfter;

  // At a block boundary — let default behavior handle (e.g. join paragraphs).
  // Run carriers fall through to the strike-through path; structural inline
  // nodes use the default delete behavior.
  if (!nodeAfter || !canCarryTrackedRunMark(nodeAfter)) {
    return false;
  }

  const hasOwnInsertion = nodeAfter.marks.some(
    (m) => m.type === insertionType && m.attrs["author"] === pluginState.author,
  );
  const hasDeletion = nodeAfter.marks.some((m) => m.type === deletionType);

  if (hasDeletion) {
    // Already deleted — skip cursor past it
    const newPos = isBackward ? rangeFrom : rangeTo;
    tr.setSelection(TextSelection.near(tr.doc.resolve(newPos)));
  } else if (hasOwnInsertion) {
    // Retract own insertion — actually delete the character
    tr.delete(rangeFrom, rangeTo);
  } else {
    // Mark as deletion instead of removing
    const delAttrs =
      findAdjacentRevisionForRange(state.doc, rangeFrom, rangeTo, "deletion", pluginState.author) ||
      makeMarkAttrs(pluginState);
    tr.addMark(rangeFrom, rangeTo, deletionType.create(delAttrs));
    // Move cursor past the deletion mark
    const newPos = isBackward ? rangeFrom : rangeTo;
    tr.setSelection(TextSelection.near(tr.doc.resolve(newPos)));
  }

  dispatch(tr.scrollIntoView());
  return true;
}

/** Apply the editor's selected-content deletion path for a host cut action. */
export const deleteSelectionAsSuggestion = (
  state: EditorState,
  dispatch: (tr: Transaction) => void,
): boolean => (state.selection.empty ? false : handleSuggestionDelete(state, dispatch, "forward"));

/**
 * Mark the text committed by an IME composition as a tracked insertion.
 *
 * Invoked from `compositionend` (deferred to a microtask) rather than from the
 * `appendTransaction` catch-all. ProseMirror commits composed text as part of
 * its own composition-end flush; adding the insertion mark *synchronously*
 * during that flush re-wraps the very text node the IME just finalized, which
 * corrupts CJK (Japanese / Chinese / Korean) input — characters duplicate or
 * garble. Running one tick later, after the composition has settled, wraps the
 * committed range safely. `markRangeAsInserted` skips nodes that already carry a
 * tracked-change mark, so this is idempotent. eigenpal/docx-editor#938.
 */
type MarkComposedAsInsertionOptions = {
  view: EditorView;
  from: number;
  pluginState: SuggestionModeState;
  replaced: Slice | null;
  compositionId: number | null;
};

function markComposedAsInsertion({
  view,
  from,
  pluginState,
  replaced,
  compositionId,
}: MarkComposedAsInsertionOptions): void {
  const insertionType = view.state.schema.marks["insertion"];
  const deletionType = view.state.schema.marks["deletion"];
  if (!insertionType || !deletionType) {
    return;
  }
  // PM leaves the cursor at the end of the committed composition.
  const to = view.state.selection.to;
  if (to <= from && !replaced) {
    return;
  }

  const tr = view.state.tr;
  tr.setMeta(SUGGESTION_META, true);
  // Keep the deferred annotation in its native composition's history event.
  // Mark-only steps have no adjacency range for history to infer this from.
  if (compositionId !== null) tr.setMeta("composition", compositionId);
  if (replaced) {
    // ProseMirror lets the browser own the composing DOM. Reinsert the original
    // selection only after that DOM has settled; changing it at compositionstart
    // makes a subsequent IME update reconcile against stale nodes and lose the
    // deletion revision.
    // Preserve the open edges: the native replacement joined the surrounding
    // paragraphs. Inserting the closed fragment splits them again around an
    // extra empty paragraph instead of restoring the selected range.
    tr.replaceRange(from, from, replaced);
    markRangeAsDeleted(
      tr,
      tr.doc,
      from,
      tr.mapping.map(from, 1),
      insertionType,
      deletionType,
      pluginState,
    );
    // Native composition keeps the starting paragraph's properties even when
    // the selection starts at its first character. Carry that committed
    // formatting to the restored end paragraph; rejection restores its own.
    const restoredEnd = tr.doc.resolve(tr.mapping.map(from, 1));
    const restoredStart = tr.doc.resolve(from);
    if (
      restoredStart.parent.type.name === "paragraph" &&
      restoredEnd.parent.type.name === "paragraph" &&
      restoredStart.parent !== restoredEnd.parent
    ) {
      const mark = expectParagraphAttrs(restoredStart.parent).pPrMark;
      if (mark?.kind === "del") {
        carryIntoParagraphLeft(
          view.state,
          tr,
          view.state.doc.resolve(from).parent,
          restoredEnd.before(),
          mark.info,
        );
      }
    }
  }
  const insertionFrom = tr.mapping.map(from, 1);
  const insertionTo = tr.mapping.map(to, 1);
  if (insertionTo <= insertionFrom) {
    if (tr.steps.length > 0) {
      view.dispatch(tr);
    }
    return;
  }
  const markAttrs =
    findAdjacentRevisionForRange(
      tr.doc,
      insertionFrom,
      insertionTo,
      "insertion",
      pluginState.author,
    ) || makeMarkAttrs(pluginState);
  markRangeAsInserted(
    tr,
    tr.doc,
    insertionFrom,
    insertionTo,
    insertionType,
    deletionType,
    markAttrs,
  );
  if (tr.steps.length === 0) {
    return;
  }

  // Collapse the caret right after the committed text, mirroring
  // applySuggestionInsert. Re-rendering the now-marked run can otherwise leave
  // the painted caret before the range; positions are stable across add-mark
  // steps but map anyway to stay correct if that changes.
  const caret = tr.mapping.map(to, 1);
  tr.setSelection(TextSelection.create(tr.doc, caret));
  view.dispatch(tr);
}

/**
 * Paragraphs inserted into one (a paste at the caret) split it: its break ends
 * the last part, which reads as the last inserted paragraph. Rejecting the
 * inserted breaks leaves that part, so it records what the paragraph read as.
 */
function recordSplitParagraph(
  tr: Transaction,
  before: PMNode | undefined,
  step: Step,
  insertedFrom: number,
  insertedTo: number,
  attrs: MarkAttrs,
): void {
  if (!(step instanceof ReplaceStep) || !before) {
    return;
  }
  // Content goes in, taking at most the paragraph's own end with it.
  const $at = before.resolve(step.from);
  if (
    $at.parent.type.name !== "paragraph" ||
    step.to > $at.after() ||
    before.textBetween(step.from, step.to) !== ""
  ) {
    return;
  }
  const split = $at.parent;
  const $first = tr.doc.resolve(insertedFrom);
  const $last = tr.doc.resolve(insertedTo);
  // The last part ends the inserted range, or, when the paste ends on a
  // closed paragraph, is the paragraph right before its end.
  let lastPos: number | null = null;
  if ($last.parent.type.name === "paragraph") {
    lastPos = $last.before();
  } else if ($last.nodeBefore?.type.name === "paragraph") {
    lastPos = insertedTo - $last.nodeBefore.nodeSize;
  }
  if ($first.parent.type.name !== "paragraph" || lastPos === null || lastPos === $first.before()) {
    return;
  }
  recordReplacedParagraphProperties({
    tr,
    position: lastPos,
    replaced: split,
    revision: { id: attrs.revisionId, author: attrs.author, date: attrs.date },
  });
}

/**
 * Create the suggestion mode plugin.
 * When active, text edits become tracked changes.
 */
export function createSuggestionModePlugin(initialActive = false, author = "User"): Plugin {
  // IME composition tracking. Each call returns a fresh Plugin, so a separate
  // view never shares this closure — the flag is per editor instance.
  //   - `composing` gates the `appendTransaction` catch-all so it never mutates
  //     the document mid-composition (mid-composition mark changes corrupt CJK
  //     input). It stays true until the deferred `compositionend` marking runs,
  //     so the catch-all also skips PM's own composition-commit transaction:
  //     `view.composing` flips false slightly before that final flush, hence the
  //     manual flag. eigenpal/docx-editor#938.
  //   - `compositionFrom` and `compositionReplaced` preserve the native
  //     replacement until compositionend can record both revisions.
  let composing = false;
  let compositionFrom: number | null = null;
  let compositionReplaced: Slice | null = null;
  let compositionId: number | null = null;

  return new Plugin({
    key: suggestionModeKey,

    state: {
      // Seed the revision-id counter from the loaded document on every
      // EditorState.create so suggesting-mode edits stay inside the signed
      // 32-bit range OOXML consumers accept (eigenpal/docx-editor#1093).
      init(_config, instance): SuggestionModeState {
        seedRevisionIdsFromDoc(instance.doc);
        return { active: initialActive, author };
      },
      apply(tr, state): SuggestionModeState {
        const nativeCompositionId = tr.getMeta("composition");
        if (composing && tr.docChanged && typeof nativeCompositionId === "number") {
          compositionId = nativeCompositionId;
        }
        const meta = tr.getMeta(suggestionModeKey);
        if (meta) {
          return { ...state, ...meta };
        }
        return state;
      },
    },

    props: {
      handleDOMEvents: {
        cut(view: EditorView, event: ClipboardEvent) {
          if (!suggestionModeKey.getState(view.state)?.active || view.state.selection.empty) {
            return false;
          }
          const data = event.clipboardData;
          if (!data) {
            return false;
          }
          const { dom, text } = view.serializeForClipboard(view.state.selection.content());
          data.clearData();
          data.setData("text/html", dom.innerHTML);
          data.setData("text/plain", text);
          event.preventDefault();
          return handleSuggestionDelete(view.state, view.dispatch, "forward");
        },
        // Remember where the composition starts and suppress the catch-all
        // while it runs. Composed text is marked as an insertion later, on
        // compositionend — never mid-composition. eigenpal/docx-editor#938.
        compositionstart(view: EditorView) {
          const pluginState = suggestionModeKey.getState(view.state);
          if (!pluginState?.active) {
            return false;
          }
          if (composing) {
            return false;
          }
          composing = true;
          compositionId = null;
          const { from, to } = view.state.selection;
          compositionFrom = from;
          compositionReplaced = from === to ? null : view.state.doc.slice(from, to);
          return false;
        },
        // Mark the committed text as a tracked insertion AFTER the composition
        // settles. PM commits the composed text in its own compositionend flush
        // (which runs after this handler); `composing` stays true across that
        // flush so the catch-all skips it, then the range is marked and the flag
        // cleared after the native final flush. See `markComposedAsInsertion`.
        compositionend(view: EditorView) {
          const pluginState = suggestionModeKey.getState(view.state);
          const from = compositionFrom;
          const replaced = compositionReplaced;
          compositionFrom = null;
          compositionReplaced = null;
          if (!pluginState?.active || from == null) {
            composing = false;
            compositionId = null;
            return false;
          }
          afterNativeCompositionFlush(() => {
            try {
              // Re-read state: suggestion mode may have been toggled off (or the
              // author changed) between scheduling and running this callback.
              const current = suggestionModeKey.getState(view.state);
              const nativeCompositionId = compositionId;
              if (current?.active) {
                markComposedAsInsertion({
                  view,
                  from,
                  pluginState: current,
                  replaced,
                  compositionId: nativeCompositionId,
                });
              }
            } finally {
              composing = false;
              compositionId = null;
            }
          });
          return false;
        },
        // Keep standalone suggestion views on the same input path as Folio's
        // runtime. The shared handler invokes handleTextInput exactly once.
        beforeinput(view, event) {
          if (!suggestionModeKey.getState(view.state)?.active || composing) return false;
          return handleEditorBeforeInput(view, event);
        },
      },
      // Intercept Enter / Backspace / Delete so paragraph-mark revisions
      // (ECMA-376 §17.13.5) get recorded instead of silently splitting or
      // joining paragraphs.
      handleKeyDown(view: EditorView, event: KeyboardEvent): boolean {
        const pluginState = suggestionModeKey.getState(view.state);
        if (!pluginState?.active) {
          return false;
        }

        // Only a plain Enter ends a paragraph: Shift-Enter breaks the line and
        // Mod-Enter breaks the page, through their own bindings.
        if (
          event.key === "Enter" &&
          !event.shiftKey &&
          !event.metaKey &&
          !event.ctrlKey &&
          !event.altKey
        ) {
          const { $from, empty } = view.state.selection;
          if (
            empty &&
            $from.parent.type.name === "paragraph" &&
            expectParagraphAttrs($from.parent).numPr?.kind === "reference"
          ) {
            return false;
          }
          return handleSuggestionEnter(view, pluginState);
        }
        // Backspace right after an autoformat puts the typed text back, as it
        // does while editing.
        if (event.key === "Backspace" && undoInputRule(view.state, view.dispatch)) {
          return true;
        }
        if (event.key === "Backspace" || event.key === "Delete") {
          const { $from, empty } = view.state.selection;
          if (
            event.key === "Backspace" &&
            empty &&
            $from.parentOffset === 0 &&
            $from.parent.type.name === "paragraph" &&
            expectParagraphAttrs($from.parent).numPr?.kind === "reference"
          ) {
            return false;
          }
          if (
            event.key === "Backspace" &&
            clearIndentOnBackspace(view.state, (tr) => {
              const revision = makeMarkAttrs(pluginState);
              recordReplacedParagraphProperties({
                tr,
                position: $from.before(),
                replaced: $from.parent,
                revision: { id: revision.revisionId, author: revision.author, date: revision.date },
              });
              view.dispatch(tr.setMeta(SUGGESTION_META, true));
            })
          ) {
            return true;
          }
          const boundaryTarget = paragraphBoundaryTarget(
            view.state,
            event.key === "Backspace" ? "backward" : "forward",
          );
          if (boundaryTarget !== null) {
            return applyPPrDel(view, boundaryTarget, pluginState);
          }
          const direction = event.key === "Backspace" ? "backward" : "forward";
          return handleSuggestionDelete(view.state, view.dispatch, direction);
        }
        return false;
      },

      // Shared beforeinput routing and native DOM reconciliation both use this hook.
      handleTextInput(view: EditorView, from: number, to: number, text: string): boolean {
        const pluginState = suggestionModeKey.getState(view.state);
        if (!pluginState?.active) {
          return false;
        }
        // During / right after an IME composition, ProseMirror has already
        // applied the composed text from the DOM. Re-inserting it here would
        // duplicate it (and desync the view), so defer to compositionend.
        // eigenpal/docx-editor#938.
        if (composing || view.composing) {
          return false;
        }
        return applySuggestionInsert(view, from, to, text, pluginState);
      },

      // Pasting over a non-empty selection must track the replaced text as a
      // deletion (not destroy it) and the pasted text as an insertion, and a
      // paste into table cells the same way in every cell it lands on. A
      // collapsed cursor falls through to the default paste + catch-all.
      handlePaste(view: EditorView, _event: ClipboardEvent, slice: Slice) {
        const pluginState = suggestionModeKey.getState(view.state);
        if (!pluginState?.active) {
          return false;
        }
        return (
          handleSuggestionTableCellPaste(view, slice, pluginState) ||
          handleSuggestionPaste(view, slice, pluginState)
        );
      },
    },

    // Catch-all: mark any unhandled new content (e.g. paste) as insertion
    appendTransaction(transactions, _oldState, newState) {
      const pluginState = suggestionModeKey.getState(newState);
      if (!pluginState?.active) {
        return null;
      }

      // Leave composed text un-marked while an IME composition is in flight.
      // `compositionend` marks the final committed range once, after the IME
      // settles; marking here (mid-composition, or during PM's compositionend
      // commit) re-wraps the active text node and corrupts CJK input.
      // eigenpal/docx-editor#938.
      if (composing) {
        return null;
      }

      const userTr = transactions.find(
        (tr) =>
          tr.docChanged &&
          !tr.getMeta(SUGGESTION_META) &&
          !tr.getMeta(SUGGESTION_BYPASS_META) &&
          !isHistoryTransaction(tr),
      );
      if (!userTr) {
        return null;
      }

      const insertionType = newState.schema.marks["insertion"];
      if (!insertionType) {
        return null;
      }

      const markAttrs = makeMarkAttrs(pluginState);

      const tr = newState.tr;
      tr.setMeta(SUGGESTION_META, true);

      const deletionType = newState.schema.marks["deletion"];
      // A step reports its range in the document right after it; the steps
      // and transactions that follow move it before `newState.doc`.
      const laterMaps = transactions
        .slice(transactions.indexOf(userTr) + 1)
        .flatMap((transaction) => transaction.mapping.maps);
      for (const [stepIndex, step] of userTr.steps.entries()) {
        const stepMap = step.getMap();
        const following = new Mapping([...userTr.mapping.maps.slice(stepIndex + 1), ...laterMaps]);
        // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror StepMap.forEach
        stepMap.forEach((_oldFrom, _oldTo, stepFrom, stepTo) => {
          // oxlint-disable-next-line unicorn/no-array-method-this-argument -- ProseMirror Mapping.map(pos, assoc)
          const newFrom = following.map(stepFrom, -1);
          // oxlint-disable-next-line unicorn/no-array-method-this-argument -- ProseMirror Mapping.map(pos, assoc)
          const newTo = following.map(stepTo, 1);
          if (newTo > newFrom) {
            if (userTr.getMeta("paste")) {
              enclosePastedRunRevisions({ tr, from: newFrom, to: newTo, revision: markAttrs });
            }
            if (deletionType) {
              markRangeAsInserted(
                tr,
                newState.doc,
                newFrom,
                newTo,
                insertionType,
                deletionType,
                markAttrs,
              );
            }
            recordSplitParagraph(tr, userTr.docs[stepIndex], step, newFrom, newTo, markAttrs);
          }
        });
      }

      return tr.steps.length > 0 ? tr : null;
    },
  });
}

/**
 * Toggle suggestion mode on/off.
 */
export function toggleSuggestionMode(
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
): boolean {
  const current = suggestionModeKey.getState(state);
  if (!current) {
    return false;
  }

  if (dispatch) {
    const tr = state.tr.setMeta(suggestionModeKey, {
      active: !current.active,
    });
    dispatch(tr);
  }
  return true;
}

/**
 * Set suggestion mode active state and author.
 */
export function setSuggestionMode(
  active: boolean,
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
  author?: string,
): boolean {
  if (dispatch) {
    const meta: Partial<SuggestionModeState> = { active };
    if (author !== undefined) {
      meta.author = author;
    }
    const tr = state.tr.setMeta(suggestionModeKey, meta);
    dispatch(tr);
  }
  return true;
}

/**
 * Check if suggestion mode is currently active.
 */
export function isSuggestionModeActive(state: EditorState): boolean {
  return suggestionModeKey.getState(state)?.active ?? false;
}
