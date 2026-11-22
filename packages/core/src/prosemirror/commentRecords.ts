/**
 * The comment threads of a ProseMirror session, held in its editor state.
 *
 * A thread's text, author, replies and resolution live beside the document,
 * not in it, yet they are part of what the user edits: deleting a comment
 * removes its record and its anchors together, and undoing that has to bring
 * both back. Keeping the records in a host list outside the editor made every
 * comment edit invisible to the editor's history, so undo restored the anchor
 * at best and the save then lost the thread.
 *
 * Here the records change only through {@link CommentRecordsStep}, a step that
 * leaves the document as it is and inverts to the records it replaced. Every
 * comment edit is a transaction holding that step and, when the anchors
 * change too, the mark and node steps beside it, so history undoes and redoes
 * a comment edit as one. Hosts mirror the records from the state; saving, the
 * controlled `comments` prop and document operations read the same records.
 */

import { panic } from "better-result";
import { closeHistory } from "prosemirror-history";
import type { Mark, Node as PMNode, Schema } from "prosemirror-model";
import {
  type EditorState,
  Plugin,
  Selection,
  type Transaction,
} from "prosemirror-state";
import { Mapping, Step, StepMap, StepResult } from "prosemirror-transform";
import type { EditorView } from "prosemirror-view";

import type { Comment } from "../types/content";
import { PENDING_COMMENT_ID } from "./commentIdAllocator";
import {
  isCommentRangeBoundaryNode,
  readCommentRangeBoundaryAttrs,
} from "./commentRangeBoundaryAttrs";
import { readCommentReferenceAttrs } from "./commentReferenceAttrs";
import { COMMENT_REFERENCE_NODE_NAME } from "./commentReferenceIntegrity";
import { RANGE_ANCHOR_NODE_NAME } from "./extensions/nodes/RangeAnchorExtension";
import {
  type DroppedRange,
  settleDroppedComments,
  settlePastedComments,
} from "./pastedCommentCopies";
import { createPluginKeyOwner } from "./pluginKeyOwner";
import { SUGGESTION_BYPASS_META } from "./plugins/suggestionMode";
import { readRangeAnchorAttrs } from "./rangeAnchorAttrs";

export type CommentRecords = readonly Comment[];

const STEP_TYPE = "folioCommentRecords";

/**
 * Replace the session's comment records. The document is unchanged, so the
 * step maps nothing; its inverse puts back the records it replaced.
 */
export class CommentRecordsStep extends Step {
  readonly before: CommentRecords;
  readonly after: CommentRecords;

  constructor(before: CommentRecords, after: CommentRecords) {
    super();
    this.before = before;
    this.after = after;
  }

  override apply(doc: PMNode): StepResult {
    return StepResult.ok(doc);
  }

  override getMap(): StepMap {
    return StepMap.empty;
  }

  override invert(): CommentRecordsStep {
    return new CommentRecordsStep(this.after, this.before);
  }

  override map(): CommentRecordsStep {
    return this;
  }

  override merge(other: Step): CommentRecordsStep | null {
    return other instanceof CommentRecordsStep && other.before === this.after
      ? new CommentRecordsStep(this.before, other.after)
      : null;
  }

  override toJSON(): unknown {
    return { stepType: STEP_TYPE, before: this.before, after: this.after };
  }

  static override fromJSON(_schema: Schema, json: unknown): CommentRecordsStep {
    if (
      typeof json !== "object" ||
      json === null ||
      !("before" in json) ||
      !("after" in json) ||
      !Array.isArray(json.before) ||
      !Array.isArray(json.after)
    ) {
      throw new RangeError("Invalid comment records step JSON");
    }
    // SAFETY: the arrays were written by `toJSON` from typed records.
    const { before, after } = json as {
      before: Comment[];
      after: Comment[];
    };
    return new CommentRecordsStep(before, after);
  }
}

Step.jsonID(STEP_TYPE, CommentRecordsStep);

/**
 * Who changed the records: `edit` for user edits and history, which hosts
 * report to their listeners; `sync` when a host wrote its own value in.
 */
export type CommentRecordsOrigin = "edit" | "sync";

type CommentRecordsPluginState = {
  records: CommentRecords;
  origin: CommentRecordsOrigin;
};

const commentRecordsOwner =
  createPluginKeyOwner<CommentRecordsPluginState>("commentRecords");

const SYNC_META = `${commentRecordsOwner.key}.sync`;

/**
 * Told the records after every change; `null` when the view's state stops
 * holding records (another session took the view over, or it was destroyed).
 */
type CommentRecordsListener = (
  records: CommentRecords | null,
  origin: CommentRecordsOrigin,
) => void;

const listeners = new WeakMap<EditorView, Set<CommentRecordsListener>>();

/** A drop's transaction, and where its content stands after it. */
type PlacedDrop = {
  type: "placed";
  root: Transaction;
  ranges: readonly DroppedRange[];
};

/** What a paste or a drop leaves for its transaction. */
type PendingPaste =
  // Its slice is settled, and these threads join the records.
  | { type: "settled"; copies: readonly Comment[] }
  // Its slice is being dragged, and is settled where it is dropped.
  | { type: "dragged" }
  | PlacedDrop;

/** Whether a transaction places a pasted or dropped slice. */
const isPlacement = (transaction: Transaction): boolean =>
  transaction.getMeta("paste") === true ||
  transaction.getMeta("uiEvent") === "drop";

const mapRanges = (
  ranges: readonly DroppedRange[],
  maps: readonly StepMap[],
): DroppedRange[] =>
  ranges.flatMap(({ from, to }) => {
    const mapping = new Mapping(maps);
    // oxlint-disable-next-line unicorn/no-array-method-this-argument -- ProseMirror Mapping.map(pos, assoc)
    const mapped = { from: mapping.map(from, 1), to: mapping.map(to, -1) };
    return mapped.from < mapped.to ? [mapped] : [];
  });

/** Where a transaction's steps put content, in its resulting document. */
const insertedRanges = ({ mapping }: Transaction): DroppedRange[] => {
  const inserted = mapping.maps.flatMap((map, index) => {
    const ranges: DroppedRange[] = [];
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror StepMap.forEach
    map.forEach((_oldFrom, _oldTo, from, to) => {
      if (from < to) ranges.push({ from, to });
    });
    return mapRanges(ranges, mapping.maps.slice(index + 1));
  });
  // A later step may put content inside what an earlier one did.
  const merged: DroppedRange[] = [];
  for (const range of inserted.toSorted((a, b) => a.from - b.from)) {
    const last = merged.at(-1);
    if (last === undefined || last.to < range.from) merged.push({ ...range });
    else last.to = Math.max(last.to, range.to);
  }
  return merged;
};

/**
 * Where a drop stands after `transactions`: the drop itself with what plugins
 * appended, or only what they appended to a drop already placed. Undefined
 * when they are not the drop's.
 */
const placedDrop = (
  paste: Exclude<PendingPaste, { type: "settled" }>,
  transactions: readonly Transaction[],
): PlacedDrop | undefined => {
  if (paste.type === "placed") {
    const { root, ranges } = paste;
    if (transactions.some((tr) => tr.getMeta("appendedTransaction") !== root))
      return undefined;
    const maps = transactions.flatMap(({ mapping }) => mapping.maps);
    return { type: "placed", root, ranges: mapRanges(ranges, maps) };
  }
  const at = transactions.findIndex(isPlacement);
  const root = transactions[at];
  if (root === undefined) return undefined;
  const maps = transactions
    .slice(at + 1)
    .flatMap(({ mapping }) => mapping.maps);
  return {
    type: "placed",
    root,
    ranges: mapRanges(insertedRanges(root), maps),
  };
};

/**
 * The plugin that holds a session's comment records, starting from `initial`
 * (the loaded document's comments).
 */
export function createCommentRecordsPlugin(
  initial: CommentRecords,
): Plugin<CommentRecordsPluginState> {
  // A comment owns one range: a paste waits here, by the state it goes into,
  // to add its threads to the records in its own undo step.
  const pending = new WeakMap<EditorState, PendingPaste>();
  return new Plugin<CommentRecordsPluginState>({
    key: commentRecordsOwner.createKey(),
    props: {
      transformPasted(slice, view, plain) {
        // Plain text holds no comment: it takes the marks at the caret.
        if (plain) return slice;
        if (view.dragging) {
          pending.set(view.state, { type: "dragged" });
          return slice;
        }
        const settled = settlePastedComments(slice, {
          doc: view.state.doc,
          records: commentRecordsOf(view.state),
          named: commentIdsNamedIn(view.state.doc),
        });
        if (settled.copies.length > 0)
          pending.set(view.state, { type: "settled", copies: settled.copies });
        return settled.slice;
      },
    },
    appendTransaction: (transactions, oldState, newState) => {
      const paste = pending.get(oldState);
      if (paste === undefined) return null;
      pending.delete(oldState);
      const records = commentRecordsOf(newState);
      if (records === null) return null;
      if (paste.type !== "settled") {
        const placed = placedDrop(paste, transactions);
        if (placed === undefined) return null;
        const transaction = newState.tr;
        const dropped = settleDroppedComments(transaction, {
          ranges: placed.ranges,
          records,
        });
        if (dropped === null) {
          // A plugin may put back what the drop removed, a tracked deletion
          // say, and the transactions that do come to this plugin next.
          pending.set(newState, placed);
          return null;
        }
        if (dropped.length > 0)
          stepCommentRecords(newState, transaction, [...records, ...dropped]);
        return transaction;
      }
      if (!transactions.some(isPlacement)) return null;
      const { copies } = paste;
      const known = new Set(records.map(({ id }) => id));
      const named = commentIdsNamedIn(newState.doc);
      const roots = new Set(
        copies.filter(({ id }) => named.has(id)).map(({ id }) => id),
      );
      const added = copies.filter(({ id }) => {
        if (known.has(id)) return false;
        let current = copies.find((comment) => comment.id === id);
        while (current?.parentId !== undefined) {
          const parent: number = current.parentId;
          current = copies.find((comment) => comment.id === parent);
        }
        return current !== undefined && roots.has(current.id);
      });
      if (added.length === 0) return null;
      const transaction = newState.tr;
      stepCommentRecords(newState, transaction, [...records, ...added]);
      return transaction;
    },
    state: {
      init: () => ({ records: initial, origin: "sync" }),
      apply: (transaction, value) => {
        let records = value.records;
        for (const step of transaction.steps) {
          if (step instanceof CommentRecordsStep) records = step.after;
        }
        if (records === value.records) return value;
        return {
          records,
          origin: transaction.getMeta(SYNC_META) === true ? "sync" : "edit",
        };
      },
    },
    view: (created) => {
      // A view given a new document's state builds this plugin's view anew.
      const current = commentRecordsOwner.getState(created.state);
      if (current) {
        for (const listener of listeners.get(created) ?? []) {
          listener(current.records, "sync");
        }
      }
      return {
        update: (view, previous) => {
          const next = commentRecordsOwner.getState(view.state);
          if (!next || next === commentRecordsOwner.getState(previous)) return;
          for (const listener of listeners.get(view) ?? []) {
            listener(next.records, next.origin);
          }
        },
        destroy: () => {
          for (const listener of listeners.get(created) ?? []) {
            listener(null, "sync");
          }
        },
      };
    },
  });
}

/** The session's comment records, or `null` when the state holds none. */
export function commentRecordsOf(state: EditorState): CommentRecords | null {
  return commentRecordsOwner.getState(state)?.records ?? null;
}

/**
 * Call `listener` whenever the view's comment records change, including a
 * state the view is given for a newly loaded document. Returns the
 * unsubscribe function.
 */
export function subscribeCommentRecords(
  view: EditorView,
  listener: CommentRecordsListener,
): () => void {
  let set = listeners.get(view);
  if (!set) {
    set = new Set();
    listeners.set(view, set);
  }
  set.add(listener);
  return () => {
    set.delete(listener);
  };
}

/** The records a transaction ends on, given the state it started from. */
export function commentRecordsAfter(
  state: EditorState,
  transaction: Transaction,
): CommentRecords | null {
  let records = commentRecordsOf(state);
  if (records === null) return null;
  for (const step of transaction.steps) {
    if (step instanceof CommentRecordsStep) records = step.after;
  }
  return records;
}

/**
 * Add the step replacing the records `transaction` ends on with `next`.
 * Returns false when the state holds no records.
 */
export function stepCommentRecords(
  state: EditorState,
  transaction: Transaction,
  next: CommentRecords,
): boolean {
  const current = commentRecordsAfter(state, transaction);
  if (current === null) return false;
  if (current !== next) transaction.step(new CommentRecordsStep(current, next));
  return true;
}

/**
 * A transaction writing a host's value into the records: kept out of the
 * undo history and reported as `sync`, so it is not echoed back to the host.
 */
export function syncCommentRecords(
  state: EditorState,
  next: CommentRecords,
): Transaction | null {
  const current = commentRecordsOf(state);
  if (current === null || current === next) return null;
  return state.tr
    .step(new CommentRecordsStep(current, next))
    .setMeta("addToHistory", false)
    .setMeta(SYNC_META, true);
}

/**
 * Whether a transaction only writes a host's comment list in: no document
 * content and no user edit changes, so it neither marks the document changed
 * nor needs an editable document.
 */
export function isCommentRecordsSync(transaction: Transaction): boolean {
  return (
    transaction.getMeta(SYNC_META) === true &&
    transaction.steps.length > 0 &&
    transaction.steps.every((step) => step instanceof CommentRecordsStep)
  );
}

// ---------------------------------------------------------------------------
// Thread edits on records
// ---------------------------------------------------------------------------

/** A comment and every reply whose parent chain reaches it. */
export function commentThreadIds(
  records: CommentRecords,
  id: number,
): Set<number> {
  const ids = new Set([id]);
  let added = true;
  while (added) {
    added = false;
    for (const comment of records) {
      if (
        !ids.has(comment.id) &&
        comment.parentId !== undefined &&
        ids.has(comment.parentId)
      ) {
        ids.add(comment.id);
        added = true;
      }
    }
  }
  return ids;
}

const editTransaction = (state: EditorState): Transaction =>
  closeHistory(state.tr).setMeta(SUGGESTION_BYPASS_META, true);

/**
 * An undoable edit replacing the records with `next`, for a host change that
 * touches no anchor (a reply, a resolution, a list the host computed).
 */
export function editCommentRecords(
  state: EditorState,
  next: CommentRecords,
): Transaction | null {
  const records = commentRecordsOf(state);
  if (records === null || records === next) return null;
  const tr = editTransaction(state);
  stepCommentRecords(state, tr, next);
  return tr;
}

/** Append comments (new threads or replies) to the records. */
export function appendCommentRecords(
  state: EditorState,
  additions: readonly Comment[],
): Transaction | null {
  const records = commentRecordsOf(state);
  if (records === null || additions.length === 0) return null;
  const tr = editTransaction(state);
  stepCommentRecords(state, tr, [...records, ...additions]);
  return tr;
}

/** Mark one comment resolved or open again. */
export function setCommentResolved(
  state: EditorState,
  id: number,
  done: boolean,
): Transaction | null {
  const records = commentRecordsOf(state);
  if (records === null) return null;
  const index = records.findIndex((comment) => comment.id === id);
  const comment = records[index];
  if (!comment || (comment.done ?? false) === done) return null;
  const next = records.with(index, { ...comment, done });
  const tr = editTransaction(state);
  stepCommentRecords(state, tr, next);
  return tr;
}

export type CommentRange = { from: number; to: number };

type DispatchingView = Pick<EditorView, "state" | "dispatch">;

const clampRange = (
  state: EditorState,
  range: CommentRange,
): CommentRange | null => {
  const size = state.doc.content.size;
  const from = Math.max(0, Math.min(range.from, size));
  const to = Math.max(from, Math.min(range.to, size));
  return from === to ? null : { from, to };
};

/**
 * Anchor a new thread on `range` and add its record, in one transaction.
 * Returns null when the range is empty or out of the document.
 */
export function addCommentThread(
  state: EditorState,
  range: CommentRange,
  comment: Comment,
): Transaction | null {
  const records = commentRecordsOf(state);
  const commentType = state.schema.marks["comment"];
  const safe = clampRange(state, range);
  if (records === null || !commentType || !safe) return null;
  const tr = editTransaction(state).addMark(
    safe.from,
    safe.to,
    commentType.create({ commentId: comment.id }),
  );
  stepCommentRecords(state, tr, [...records, comment]);
  return tr;
}

/**
 * Show the add-comment flow's highlight on `range`. The highlight is not an
 * edit: it stays out of the undo history, so undoing a comment never brings
 * it back.
 */
export function markPendingComment(
  view: DispatchingView,
  range: CommentRange,
  options?: { selectEnd?: boolean },
): boolean {
  const commentType = view.state.schema.marks["comment"];
  const safe = clampRange(view.state, range);
  if (!commentType || !safe) return false;
  const tr = view.state.tr
    .addMark(
      safe.from,
      safe.to,
      commentType.create({ commentId: PENDING_COMMENT_ID }),
    )
    .setMeta("addToHistory", false)
    .setMeta(SUGGESTION_BYPASS_META, true);
  if (options?.selectEnd) {
    tr.setSelection(Selection.near(tr.doc.resolve(safe.to), -1));
  }
  view.dispatch(tr);
  return true;
}

/** Remove the add-comment flow's highlight from `range`. */
export function clearPendingComment(
  view: DispatchingView,
  range: CommentRange,
): void {
  const commentType = view.state.schema.marks["comment"];
  const safe = clampRange(view.state, range);
  if (!commentType || !safe) return;
  const pending = commentType.create({ commentId: PENDING_COMMENT_ID });
  if (!view.state.doc.rangeHasMark(safe.from, safe.to, pending)) return;
  view.dispatch(
    view.state.tr
      .removeMark(safe.from, safe.to, pending)
      .setMeta("addToHistory", false)
      .setMeta(SUGGESTION_BYPASS_META, true),
  );
}

/**
 * The add-comment flow's commit: the highlight on `range` gives way to the
 * new thread's mark, and the thread joins the records, as one undoable edit.
 */
export function commitCommentThread(
  view: DispatchingView,
  range: CommentRange,
  comment: Comment,
): boolean {
  if (!addCommentThread(view.state, range, comment)) return false;
  clearPendingComment(view, range);
  const tr = addCommentThread(view.state, range, comment);
  if (!tr) return false;
  view.dispatch(tr);
  return true;
}

const anchoredCommentId = (node: PMNode): number | undefined => {
  if (node.type.name === COMMENT_REFERENCE_NODE_NAME) {
    const attrs = readCommentReferenceAttrs(node);
    return attrs.ok ? attrs.value.commentId : undefined;
  }
  if (isCommentRangeBoundaryNode(node)) {
    const attrs = readCommentRangeBoundaryAttrs(node);
    return attrs.ok ? attrs.value.marker.id : undefined;
  }
  return undefined;
};

/**
 * Remove every anchor of the comments in `ids` from the transaction's
 * document: their `comment` marks, references and range boundaries.
 */
export function removeCommentAnchors(
  transaction: Transaction,
  ids: ReadonlySet<number>,
): void {
  const marks: { from: number; to: number; mark: Mark }[] = [];
  const nodes: (
    | { type: "delete"; from: number; to: number }
    | { type: "replace"; from: number; to: number; replacement: PMNode }
  )[] = [];
  transaction.doc.descendants((node, position) => {
    for (const mark of node.marks) {
      const commentId: unknown = mark.attrs["commentId"];
      if (
        mark.type.name === "comment" &&
        typeof commentId === "number" &&
        ids.has(commentId)
      ) {
        marks.push({ from: position, to: position + node.nodeSize, mark });
      }
    }
    if (node.type.name === RANGE_ANCHOR_NODE_NAME) {
      const attrs = readRangeAnchorAttrs(node);
      if (!attrs.ok) return false;
      const markers = attrs.value.markers.filter(
        (marker) =>
          (marker.type !== "commentRangeStart" &&
            marker.type !== "commentRangeEnd") ||
          !ids.has(marker.id),
      );
      if (markers.length !== attrs.value.markers.length)
        nodes.push(
          markers.length === 0
            ? { type: "delete", from: position, to: position + node.nodeSize }
            : {
                type: "replace",
                from: position,
                to: position + node.nodeSize,
                replacement: node.type.create({ markers }, null, node.marks),
              },
        );
      return false;
    }
    const anchored = anchoredCommentId(node);
    if (anchored !== undefined && ids.has(anchored)) {
      nodes.push({
        type: "delete",
        from: position,
        to: position + node.nodeSize,
      });
      return false;
    }
    return true;
  });
  for (const { from, to, mark } of marks) {
    transaction.removeMark(from, to, mark);
  }
  // Last first, so the earlier positions stay valid.
  for (const node of nodes.toReversed()) {
    switch (node.type) {
      case "delete":
        transaction.delete(node.from, node.to);
        break;
      case "replace":
        transaction.replaceWith(node.from, node.to, node.replacement);
        break;
      default: {
        const exhaustive: never = node;
        panic(`Unknown comment anchor change: ${exhaustive}`);
      }
    }
  }
}

/** Delete a thread: its record, every reply under it, and its anchors. */
export function deleteCommentThread(
  state: EditorState,
  id: number,
): Transaction | null {
  const records = commentRecordsOf(state);
  if (records === null) return null;
  const ids = commentThreadIds(records, id);
  const next = records.filter((comment) => !ids.has(comment.id));
  const tr = editTransaction(state);
  removeCommentAnchors(tr, ids);
  stepCommentRecords(state, tr, next);
  return tr.steps.length > 0 ? tr : null;
}

/**
 * The comment ids a document names through a mark, reference or range
 * boundary.
 */
export function commentIdsNamedIn(doc: PMNode): Set<number> {
  const ids = new Set<number>();
  doc.descendants((node) => {
    for (const mark of node.marks) {
      const commentId: unknown = mark.attrs["commentId"];
      if (mark.type.name === "comment" && typeof commentId === "number")
        ids.add(commentId);
    }
    if (node.type.name === RANGE_ANCHOR_NODE_NAME) {
      const attrs = readRangeAnchorAttrs(node);
      if (attrs.ok)
        for (const marker of attrs.value.markers)
          if (marker.type === "commentRangeStart") ids.add(marker.id);
    }
    const anchored = anchoredCommentId(node);
    if (anchored !== undefined) ids.add(anchored);
    return true;
  });
  return ids;
}

/**
 * Remove the threads a review resolution left without any anchor: accepting
 * the deletion of everything a comment covers, or rejecting the suggestion
 * that inserted it, retires the comment with its replies, in the same
 * transaction. Other edits keep records whose anchors are gone for now (a
 * cut keeps the thread for the paste).
 */
export function retireThreadsResolvedAway(
  state: EditorState,
  transaction: Transaction,
): void {
  if (transaction.doc === state.doc) return;
  const records = commentRecordsAfter(state, transaction);
  if (records === null || records.length === 0) return;
  const before = commentIdsNamedIn(state.doc);
  if (before.size === 0) return;
  const after = commentIdsNamedIn(transaction.doc);
  const retired = new Set<number>();
  for (const id of before) {
    if (after.has(id) || id === PENDING_COMMENT_ID) continue;
    for (const threadId of commentThreadIds(records, id)) retired.add(threadId);
  }
  if (retired.size === 0) return;
  const next = records.filter(({ id }) => !retired.has(id));
  if (next.length !== records.length)
    stepCommentRecords(state, transaction, next);
}

/**
 * Add the records of `created` comments that `transaction`'s document names
 * and the records lack, so comments a document operation anchors join the
 * records in the operation's own transaction.
 */
export function recordCreatedComments(
  state: EditorState,
  transaction: Transaction,
  created: readonly Comment[],
): void {
  if (created.length === 0) return;
  const records = commentRecordsAfter(state, transaction);
  if (records === null) return;
  const known = new Set(records.map(({ id }) => id));
  const named = commentIdsNamedIn(transaction.doc);
  const additions = created.filter(({ id }) => named.has(id) && !known.has(id));
  if (additions.length > 0)
    stepCommentRecords(state, transaction, [...records, ...additions]);
}

/**
 * `view` for a document operation: its dispatch adds the records of the
 * `created` comments the operation anchors, in the operation's transaction.
 */
export function recordingCommentsView(
  view: DispatchingView,
  created: readonly Comment[],
): DispatchingView {
  return {
    get state() {
      return view.state;
    },
    dispatch: (transaction) => {
      recordCreatedComments(view.state, transaction, created);
      view.dispatch(transaction);
    },
  };
}

/** Panic unless the state holds records; for callers that require them. */
export function expectCommentRecords(state: EditorState): CommentRecords {
  return (
    commentRecordsOf(state) ?? panic("The editor state holds no comments.")
  );
}
