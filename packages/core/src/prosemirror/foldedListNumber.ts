/**
 * The editor's side of the list-number fold (`docx/foldedListNumberFields`).
 *
 * A capture of a folded `LISTNUM` field is an atom that shows nothing, which
 * is right only while the paragraph's marker shows the field for it. Edits
 * can end that: the paragraph loses its numbering or changes level, it is
 * joined behind another, the capture is moved or pasted. So after every
 * change the paragraphs are brought back to the one form the fold allows, and
 * a capture that no marker shows becomes the ordinary field, or tab, it stood
 * for. What the page shows is then what a save writes.
 */

import { paragraphNumberingReferenceId } from "@stll/docx-core/model";
import { Fragment, type Mark, type Node as PMNode, Slice } from "prosemirror-model";
import { type EditorState, Plugin, PluginKey, type Transaction } from "prosemirror-state";
import {
  AddMarkStep,
  AttrStep,
  RemoveMarkStep,
  ReplaceAroundStep,
  ReplaceStep,
  type Step,
} from "prosemirror-transform";

import {
  cachedListNumberText,
  isFoldedListNumber,
  type ListMarkerFoldState,
  listMarkerFoldState,
  listMarkerWithFold,
  type ListNumberFoldItem,
  planListNumberFold,
} from "../docx/foldedListNumberFields";
import type { FoldedListNumber } from "../types/document";
import { SUGGESTION_BYPASS_META } from "./plugins/suggestionMode";
import { type PositionQuery, sweepPositions } from "./positionSweep";
import { RUN_IDENTITY_MARK_NAME } from "./runIdentity";

const PRESERVED_XML = "preservedXml";

/** Atoms that show nothing and so do not end the start of a paragraph. */
const HIDDEN_NODE_NAMES: ReadonlySet<string> = new Set([
  "bookmarkBoundary",
  "moveRangeBoundary",
  "rangeAnchor",
]);

/** What a capture node stands for, or nothing for any other node. */
export const foldedListNumberOfNode = (node: PMNode): FoldedListNumber | undefined => {
  if (node.type.name !== PRESERVED_XML) {
    return undefined;
  }
  const folded: unknown = node.attrs["foldedListNumber"];
  return isFoldedListNumber(folded) ? folded : undefined;
};

/** Marks under which content is written as a tracked change. */
const TRACKED_MARK_NAMES: ReadonlySet<string> = new Set(["insertion", "deletion"]);

const isTracked = (node: PMNode): boolean =>
  node.marks.some((mark) => TRACKED_MARK_NAMES.has(mark.type.name));

const foldItemOf = (node: PMNode): ListNumberFoldItem => {
  const folded = foldedListNumberOfNode(node);
  // A capture under a tracked change is never hidden: the change is written
  // with spellings of its own, which the field has and its markup does not.
  if (folded && isTracked(node)) {
    return { kind: "shown" };
  }
  if (folded) {
    return folded.kind === "field"
      ? { kind: "field", cached: cachedListNumberText(folded.field) }
      : { kind: "tab" };
  }
  if (HIDDEN_NODE_NAMES.has(node.type.name)) {
    return { kind: "hidden" };
  }
  return node.type.name === PRESERVED_XML && node.attrs["text"] === ""
    ? { kind: "hidden" }
    : { kind: "shown" };
};

/**
 * `node` under `marks`, and without a run identity: it is new content where it
 * lands, and a run it shares an id with by accident is not its run.
 */
const withMarks = (node: PMNode, marks: readonly Mark[]): PMNode => {
  let set: readonly Mark[] = node.marks.filter((mark) => mark.type.name !== RUN_IDENTITY_MARK_NAME);
  for (const mark of marks) {
    set = mark.addToSet(set);
  }
  return node.mark(set);
};

/**
 * The nodes that show what `capture` stands for: the field as the editor
 * shows any field, the tab as a tab, under the marks the capture carried.
 * The capture itself when it carries nothing to show in its place.
 */
export const unfoldedListNumberNodes = (capture: PMNode): PMNode[] => {
  const json: unknown = capture.attrs["foldedListNumberNodes"];
  if (foldedListNumberOfNode(capture) === undefined || !Array.isArray(json) || json.length === 0) {
    return [capture];
  }
  try {
    return json.map((node) => withMarks(capture.type.schema.nodeFromJSON(node), capture.marks));
  } catch {
    return [capture];
  }
};

type Child = { node: PMNode; from: number };

const paragraphFoldState = (paragraph: PMNode): ListMarkerFoldState => {
  const marker: unknown = paragraph.attrs["listMarker"];
  const template: unknown = paragraph.attrs["listMarkerTemplate"];
  // Most paragraphs: a marker with no tab shows no folded field.
  if (typeof marker !== "string" || !marker.includes("\t")) {
    return { showsFields: false, base: typeof marker === "string" ? marker : "" };
  }
  return listMarkerFoldState({
    marker,
    template: typeof template === "string" ? template : null,
    isBullet: paragraph.attrs["listIsBullet"] === true,
    numbered: paragraphNumberingReferenceId(paragraph.attrs["numPr"] ?? undefined) !== undefined,
  });
};

const holdsCapture = (paragraph: PMNode): boolean => {
  for (let index = 0; index < paragraph.childCount; index += 1) {
    if (foldedListNumberOfNode(paragraph.child(index)) !== undefined) {
      return true;
    }
  }
  return false;
};

/** Bring one paragraph to the form the fold allows. Positions are of the document `tr` began with. */
const normalizeParagraph = (tr: Transaction, paragraph: PMNode, position: number): void => {
  const state = paragraphFoldState(paragraph);
  if (!state.showsFields && !holdsCapture(paragraph)) {
    return;
  }
  const children: Child[] = [];
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  paragraph.forEach((node, offset) => {
    children.push({ node, from: position + 1 + offset });
  });

  const plan = planListNumberFold(
    children.map(({ node }) => foldItemOf(node)),
    state.showsFields,
  );

  if (state.showsFields) {
    const marker = listMarkerWithFold(state, plan.suffix);
    if (marker !== paragraph.attrs["listMarker"]) {
      tr.setNodeAttribute(tr.mapping.map(position), "listMarker", marker);
    }
  }

  // From the last to the first, so no replacement moves one still to make.
  for (let index = children.length - 1; index >= 0; index -= 1) {
    const child = children[index];
    if (!child || foldedListNumberOfNode(child.node) === undefined || plan.hidden.has(index)) {
      continue;
    }
    const shown = unfoldedListNumberNodes(child.node);
    if (shown.length === 1 && shown[0] === child.node) {
      continue;
    }
    tr.replaceWith(
      tr.mapping.map(child.from, -1),
      tr.mapping.map(child.from + child.node.nodeSize),
      shown,
    );
  }

  // The plan moves one stretch, to the start: what it puts first that was not first.
  const first = plan.order[0];
  if (first === undefined || first === 0) {
    return;
  }
  let moved = 1;
  while (plan.order[moved] === first + moved) {
    moved += 1;
  }
  const head = children[first];
  const tail = children[first + moved - 1];
  if (!head || !tail) {
    return;
  }
  const from = tr.mapping.map(head.from, -1);
  const to = tr.mapping.map(tail.from + tail.node.nodeSize);
  const stretch = tr.doc.slice(from, to).content;
  tr.delete(from, to);
  tr.insert(tr.mapping.map(position + 1, -1), stretch);
};

type Range = { from: number; to: number };

/** A paragraph the fold has something to say about: it holds a capture, or its marker shows one. */
const concernsFold = (paragraph: PMNode): boolean => {
  const marker: unknown = paragraph.attrs["listMarker"];
  return (typeof marker === "string" && marker.includes("\t")) || holdsCapture(paragraph);
};

const fragmentConcernsFold = (fragment: Fragment): boolean => {
  let found = false;
  fragment.descendants((node) => {
    if (found) {
      return false;
    }
    if (node.type.name === "paragraph") {
      found = concernsFold(node);
    } else if (foldedListNumberOfNode(node) !== undefined) {
      found = true;
    }
    return !found;
  });
  return found;
};

/**
 * Whether a step can bring the fold something to do into a document that had
 * nothing: a capture, or a paragraph whose marker shows one. Only what the
 * step inserts is read, never its map.
 */
const stepConcernsFold = (step: Step): boolean => {
  if (step instanceof ReplaceStep || step instanceof ReplaceAroundStep) {
    return fragmentConcernsFold(step.slice.content);
  }
  if (step instanceof AttrStep) {
    return (
      step.attr === "foldedListNumber" ||
      (step.attr === "listMarker" && typeof step.value === "string" && step.value.includes("\t"))
    );
  }
  return false;
};

const concerning = new WeakMap<Transaction, boolean>();

const transactionConcernsFold = (transaction: Transaction): boolean => {
  const known = concerning.get(transaction);
  if (known !== undefined) {
    return known;
  }
  const found = transaction.steps.some(stepConcernsFold);
  concerning.set(transaction, found);
  return found;
};

/** Sorted ranges with the overlapping and the touching ones made one. */
const coalesced = (ranges: readonly Range[]): Range[] => {
  const merged: Range[] = [];
  for (const range of ranges.toSorted((a, b) => a.from - b.from)) {
    const last = merged.at(-1);
    if (last && range.from <= last.to) {
      last.to = Math.max(last.to, range.to);
    } else {
      merged.push({ ...range });
    }
  }
  return merged;
};

/** Whether `range` meets any of `sorted`, which do not overlap. */
const meets = (sorted: readonly Range[], range: Range): boolean => {
  let low = 0;
  let high = sorted.length - 1;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    const candidate = sorted[middle];
    if (!candidate || candidate.to < range.from) {
      low = middle + 1;
    } else if (candidate.from > range.to) {
      high = middle - 1;
    } else {
      return true;
    }
  }
  return false;
};

type Reach = {
  /** Where to look, in the document the transactions left: coalesced, in order. */
  ranges: Range[];
  /** Paragraphs known to concern the fold that no change came near, where they now start. */
  carried: number[];
};

/**
 * Where `transactions` can have changed what the fold decides, given the
 * paragraphs that concerned it before them (`known`, by start, in `before`).
 *
 * Those are the known paragraphs a change touched, with the room they had so
 * that both halves of a split and the whole of a join are covered, and
 * whatever a step that inserts a capture or a marker put in. Every position is
 * carried through the step maps in one sweep, so the cost grows with the
 * steps plus the known paragraphs, not with their product.
 */
const reachOf = (
  before: PMNode,
  known: readonly number[],
  transactions: readonly Transaction[],
): Reach => {
  const queries: PositionQuery[] = [];
  // Two queries per changed range, then two per known paragraph.
  const inserts: boolean[] = [];
  let offset = 0;
  for (const transaction of transactions) {
    for (const [index, step] of transaction.steps.entries()) {
      const from = offset + index + 1;
      const concerns = stepConcernsFold(step);
      if (step instanceof AttrStep) {
        queries.push({ pos: step.pos, assoc: -1, from }, { pos: step.pos + 1, assoc: 1, from });
        inserts.push(concerns);
        continue;
      }
      // A mark moves nothing, but a tracked one decides whether a capture may stay hidden.
      if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
        if (TRACKED_MARK_NAMES.has(step.mark.type.name)) {
          queries.push({ pos: step.from, assoc: -1, from }, { pos: step.to, assoc: 1, from });
          inserts.push(false);
        }
        continue;
      }
      step.getMap().forEach((_oldStart, _oldEnd, newStart, newEnd) => {
        queries.push({ pos: newStart, assoc: -1, from }, { pos: newEnd, assoc: 1, from });
        inserts.push(concerns);
      });
    }
    offset += transaction.steps.length;
  }
  const changes = inserts.length;
  for (const start of known) {
    const paragraph = before.nodeAt(start);
    const end = start + (paragraph?.nodeSize ?? 0);
    queries.push({ pos: start, assoc: -1, from: 0 }, { pos: end, assoc: 1, from: 0 });
  }

  const swept = sweepPositions(
    transactions.map((transaction) => transaction.mapping),
    queries,
  );
  const rangeAt = (index: number): Range => ({
    from: swept[index * 2]?.pos ?? 0,
    to: swept[index * 2 + 1]?.pos ?? 0,
  });

  const changed: Range[] = [];
  const ranges: Range[] = [];
  for (let index = 0; index < changes; index += 1) {
    const range = rangeAt(index);
    changed.push(range);
    if (inserts[index]) {
      ranges.push(range);
    }
  }
  const touched = coalesced(changed);
  const carried: number[] = [];
  for (let index = 0; index < known.length; index += 1) {
    const extent = rangeAt(changes + index);
    if (meets(touched, extent)) {
      ranges.push(extent);
    } else {
      carried.push(extent.from);
    }
  }
  return { ranges: coalesced(ranges), carried };
};

/** The paragraphs standing in or beside `ranges`, by start, in order. */
const paragraphsIn = (doc: PMNode, ranges: readonly Range[]): Map<number, PMNode> => {
  const paragraphs = new Map<number, PMNode>();
  const size = doc.content.size;
  for (const { from, to } of ranges) {
    const start = Math.max(0, Math.min(from, size) - 1);
    const end = Math.min(size, Math.max(to, from) + 1);
    doc.nodesBetween(start, end, (node, position) => {
      if (node.type.name === "paragraph") {
        paragraphs.set(position, node);
      }
      return true;
    });
  }
  return paragraphs;
};

const everyParagraph = (doc: PMNode): Map<number, PMNode> => {
  const paragraphs = new Map<number, PMNode>();
  doc.descendants((node, position) => {
    if (node.type.name === "paragraph") {
      paragraphs.set(position, node);
    }
    return true;
  });
  return paragraphs;
};

/** The paragraphs of the document that concern the fold, by start, in order. */
type FoldState = { paragraphs: readonly number[] };

const concerningStarts = (paragraphs: ReadonlyMap<number, PMNode>): number[] => {
  const starts: number[] = [];
  for (const [position, node] of paragraphs) {
    if (concernsFold(node)) {
      starts.push(position);
    }
  }
  return starts;
};

const foldedListNumberKey = new PluginKey<FoldState>("foldedListNumber");

export type FoldedListNumberPassOptions = {
  /** Called once for every paragraph the pass looks at. */
  onParagraphVisit?: () => void;
};

const normalizeParagraphs = (
  state: EditorState,
  paragraphs: ReadonlyMap<number, PMNode>,
  { onParagraphVisit }: FoldedListNumberPassOptions,
): Transaction | null => {
  const tr = state.tr;
  for (const position of [...paragraphs.keys()].toSorted((a, b) => a - b)) {
    const node = paragraphs.get(position);
    if (node) {
      onParagraphVisit?.();
      normalizeParagraph(tr, node, position);
    }
  }
  // Not an edit of the user's: nothing here is a tracked change.
  return tr.docChanged ? tr.setMeta(SUGGESTION_BYPASS_META, true) : null;
};

/**
 * A transaction that brings every paragraph of `state` to the form the fold
 * allows, or nothing when they already have it.
 */
export const normalizeFoldedListNumbers = (
  state: EditorState,
  options: FoldedListNumberPassOptions = {},
): Transaction | null => normalizeParagraphs(state, everyParagraph(state.doc), options);

/**
 * Keeps every capture hidden only where a marker shows it, after each change.
 *
 * Its state is the paragraphs that concern the fold at all. While there are
 * none and a transaction inserts none, the pass does nothing and reads no
 * step map. Otherwise it looks only at those paragraphs a change touched and
 * at what the transaction inserted.
 */
export const foldedListNumberPlugin = (options: FoldedListNumberPassOptions = {}): Plugin =>
  new Plugin<FoldState>({
    key: foldedListNumberKey,
    state: {
      init: (_config, state) => ({
        paragraphs: concerningStarts(everyParagraph(state.doc)),
      }),
      apply(transaction, value) {
        if (!transaction.docChanged) {
          return value;
        }
        if (value.paragraphs.length === 0 && !transactionConcernsFold(transaction)) {
          return value;
        }
        const reach = reachOf(transaction.before, value.paragraphs, [transaction]);
        const found = concerningStarts(paragraphsIn(transaction.doc, reach.ranges));
        return {
          paragraphs: [...new Set([...reach.carried, ...found])].toSorted((a, b) => a - b),
        };
      },
    },
    appendTransaction(transactions, oldState, newState) {
      const changes = transactions.filter(({ docChanged }) => docChanged);
      if (changes.length === 0) {
        return null;
      }
      const known = foldedListNumberKey.getState(oldState)?.paragraphs ?? [];
      if (known.length === 0 && !changes.some(transactionConcernsFold)) {
        return null;
      }
      const reach = reachOf(oldState.doc, known, transactions);
      return normalizeParagraphs(newState, paragraphsIn(newState.doc, reach.ranges), options);
    },
  });

/**
 * Put the folded fields of a pasted slice on the line.
 *
 * A paste cannot tell a copy from a move, and a marker somewhere else may
 * still show the field. A field on the line is one the reader can see and
 * remove; a second hidden one is neither.
 */
export const unfoldPastedListNumberFields = (slice: Slice): Slice => {
  let found = false;
  slice.content.descendants((node) => {
    if (foldedListNumberOfNode(node) !== undefined) {
      found = true;
    }
    return !found;
  });
  if (!found) {
    return slice;
  }

  const unfold = (fragment: Fragment): Fragment => {
    const children: PMNode[] = [];
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Fragment.forEach
    fragment.forEach((node) => {
      if (foldedListNumberOfNode(node) !== undefined) {
        children.push(...unfoldedListNumberNodes(node));
        return;
      }
      children.push(node.childCount === 0 ? node : node.copy(unfold(node.content)));
    });
    return Fragment.fromArray(children);
  };

  return new Slice(unfold(slice.content), slice.openStart, slice.openEnd);
};
