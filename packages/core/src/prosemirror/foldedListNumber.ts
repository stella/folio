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
import { type EditorState, Plugin, type Transaction } from "prosemirror-state";
import { AttrStep } from "prosemirror-transform";

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

const foldItemOf = (node: PMNode): ListNumberFoldItem => {
  const folded = foldedListNumberOfNode(node);
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

/**
 * Where `transactions` changed the document, in the positions of the document
 * they left. A step that replaces content says so in its map; a step that
 * sets one node's attribute has an empty map and says where in `pos`. Steps
 * that only add or remove marks change neither content nor a marker.
 */
const changedRanges = (transactions: readonly Transaction[]): Range[] => {
  const ranges: Range[] = [];
  for (const [index, transaction] of transactions.entries()) {
    const later = transactions.slice(index + 1);
    const toFinal = (position: number, assoc: number, step: number): number => {
      let mapped = transaction.mapping.slice(step + 1).map(position, assoc);
      for (const after of later) {
        mapped = after.mapping.map(mapped, assoc);
      }
      return mapped;
    };
    for (const [step, made] of transaction.steps.entries()) {
      if (made instanceof AttrStep) {
        const at = toFinal(made.pos, 1, step);
        ranges.push({ from: at, to: at + 1 });
        continue;
      }
      made.getMap().forEach((_oldStart, _oldEnd, newStart, newEnd) => {
        ranges.push({ from: toFinal(newStart, -1, step), to: toFinal(newEnd, 1, step) });
      });
    }
  }
  return ranges;
};

/**
 * The paragraphs a change can have taken out of the form the fold allows:
 * those the changed ranges touch, and the one on either side of each range,
 * since a join or a split changes the paragraph next to where it lands.
 */
const paragraphsNear = (doc: PMNode, ranges: readonly Range[]): Map<number, PMNode> => {
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

/**
 * A transaction that brings the paragraphs of `state` to the form the fold
 * allows, or nothing when they already have it. With `transactions`, only the
 * paragraphs they changed, and the neighbours of each change, are looked at;
 * without, every paragraph is.
 */
export const normalizeFoldedListNumbers = (
  state: EditorState,
  transactions?: readonly Transaction[],
): Transaction | null => {
  const paragraphs = new Map<number, PMNode>();
  if (transactions) {
    for (const [position, node] of paragraphsNear(state.doc, changedRanges(transactions))) {
      paragraphs.set(position, node);
    }
  } else {
    state.doc.descendants((node, position) => {
      if (node.type.name === "paragraph") {
        paragraphs.set(position, node);
      }
      return true;
    });
  }

  const tr = state.tr;
  for (const position of [...paragraphs.keys()].toSorted((a, b) => a - b)) {
    const node = paragraphs.get(position);
    if (node) {
      normalizeParagraph(tr, node, position);
    }
  }
  // Not an edit of the user's: nothing here is a tracked change.
  return tr.docChanged ? tr.setMeta(SUGGESTION_BYPASS_META, true) : null;
};

/** Keeps every capture hidden only where a marker shows it, after each change. */
export const foldedListNumberPlugin = (): Plugin =>
  new Plugin({
    appendTransaction(transactions, _oldState, newState) {
      const changes = transactions.filter(({ docChanged }) => docChanged);
      if (changes.length === 0) {
        return null;
      }
      return normalizeFoldedListNumbers(newState, transactions);
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
