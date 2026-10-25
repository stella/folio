/**
 * The projection of a paragraph that holds nothing but text boxes.
 *
 * Opening a package folds such a paragraph into its boxes: the parser keeps
 * the wrapper only when it carries a section break, a property change, a
 * bookmark or an empty hyperlink, and the first box then stands in for it
 * (`_docxPlacement: "standalone"`). A document edited in place can leave a
 * paragraph in that state without having been opened that way, for instance
 * once its text is rejected. Folding it the same way keeps the edited
 * document and the package it saves to one projection, positions included.
 */

import { Fragment, type Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";

import {
  anchoredTextBoxId,
  endOfParagraphWithItsTextBoxes,
  textBoxAnchorIdOf,
} from "./anchoredTextBoxes";
import { expectTextBoxAttrs } from "./attrs";
import { hostParagraphAttrs } from "./conversion/toProseDoc";

const carriesValue = (value: unknown): boolean =>
  Array.isArray(value)
    ? value.length > 0
    : value !== null && value !== undefined;

/** Whether `node` is a paragraph of text box anchors that the parser would not keep. */
const isEmptyTextBoxHost = (node: PMNode): boolean => {
  if (node.type.name !== "paragraph" || node.childCount === 0) return false;
  let onlyAnchors = true;
  node.forEach((child) => {
    if (textBoxAnchorIdOf(child) === undefined) onlyAnchors = false;
  });
  if (!onlyAnchors) return false;
  return !(
    carriesValue(node.attrs["_sectionProperties"]) ||
    carriesValue(node.attrs["_propertyChanges"]) ||
    carriesValue(node.attrs["bookmarks"]) ||
    carriesValue(node.attrs["_emptyHyperlinks"]) ||
    carriesValue(node.attrs["pPrMark"]) ||
    carriesValue(node.attrs["_suggestedInsert"]) ||
    carriesValue(node.attrs["reviewCarrier"])
  );
};

/** The host folded into the boxes drawn in it, or `undefined` when it is not whole. */
const foldedBoxes = (
  host: PMNode,
  boxes: readonly PMNode[],
): PMNode[] | undefined => {
  const anchors = new Set<string>();
  host.forEach((child) => {
    const anchorId = textBoxAnchorIdOf(child);
    if (anchorId !== undefined) anchors.add(anchorId);
  });
  const ids = boxes.map((box) => anchoredTextBoxId(box));
  if (
    boxes.length === 0 ||
    ids.some((id) => id === undefined || !anchors.has(id)) ||
    new Set(ids).size !== anchors.size
  ) {
    return undefined;
  }
  return boxes.map((box, index) => {
    const attrs = expectTextBoxAttrs(box);
    const preserved = host.attrs["_preservedAttributes"];
    return box.type.create(
      {
        ...box.attrs,
        _docxPlacement: "standalone",
        ...(index === 0
          ? {
              _docxHostParagraph: hostParagraphAttrs(host.attrs),
              ...(Array.isArray(preserved) && preserved.length > 0
                ? { _preservedAttributes: preserved }
                : { _preservedAttributes: attrs._preservedAttributes }),
            }
          : {}),
      },
      box.content,
      box.marks,
    );
  });
};

/**
 * `state` with every paragraph that hosts only text boxes folded into them,
 * or `state` itself when there is none.
 */
export const collapseEmptyTextBoxHosts = (state: EditorState): EditorState => {
  const hosts: { from: number; to: number; boxes: PMNode[] }[] = [];
  state.doc.descendants((node, position) => {
    if (!isEmptyTextBoxHost(node)) return !node.isTextblock;
    const end = endOfParagraphWithItsTextBoxes(state.doc, position);
    const boxes: PMNode[] = [];
    state.doc.nodesBetween(position + node.nodeSize, end, (child, at) => {
      if (at >= position + node.nodeSize && child.type.name === "textBox") {
        boxes.push(child);
      }
      return false;
    });
    const folded = foldedBoxes(node, boxes);
    if (folded) hosts.push({ from: position, to: end, boxes: folded });
    return false;
  });
  if (hosts.length === 0) return state;
  const tr = state.tr;
  for (const { from, to, boxes } of hosts.toReversed()) {
    tr.replaceWith(from, to, Fragment.fromArray(boxes));
  }
  tr.setMeta("addToHistory", false);
  return state.apply(tr);
};
