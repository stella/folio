import type { Node as PMNode } from "prosemirror-model";

import { subtreeHolds } from "./subtreeHolds";

export type PageBreakRunDescendantIndex = {
  firstPageBreakRunPosition: (node: PMNode) => number | undefined;
};

const indexesPageBreakDescendants = (node: PMNode): boolean => {
  switch (node.type.name) {
    case "paragraph":
    case "tableCell":
    case "tableHeader":
    case "textBox":
      return true;
    default:
      return false;
  }
};

const holdsPageBreakRun = subtreeHolds(
  (node) => node.type.name === "pageBreakRun",
);

const buildIndex = (
  doc: PMNode,
  skip: (node: PMNode) => boolean,
): PageBreakRunDescendantIndex => {
  const positions = new WeakMap<PMNode, number>();

  const visitNode = (node: PMNode, startPos: number): number | undefined => {
    let firstPosition =
      node.type.name === "pageBreakRun" ? startPos : undefined;

    // Visit every child even after finding the first break. Ancestor answers
    // stay first-in-document-order while each indexed descendant gets its own answer.
    node.forEach((child, offset) => {
      if (skip(child)) return;
      const descendantPosition = visitNode(child, startPos + 1 + offset);
      firstPosition ??= descendantPosition;
    });

    if (firstPosition !== undefined && indexesPageBreakDescendants(node)) {
      positions.set(node, firstPosition);
    }
    return firstPosition;
  };

  visitNode(doc, -1);
  return {
    firstPageBreakRunPosition: (node) => positions.get(node),
  };
};

/**
 * Build one identity-scoped index for a single layout conversion. Subtrees
 * without a page break run hold no answer, so the walk skips them.
 */
export const buildPageBreakRunDescendantIndex = (
  doc: PMNode,
): PageBreakRunDescendantIndex =>
  buildIndex(doc, (node) => !holdsPageBreakRun(node));

/** The index built visiting every node, skipping none: the reference for the skip. */
export const buildPageBreakRunDescendantIndexVisitingEveryNode = (
  doc: PMNode,
): PageBreakRunDescendantIndex => buildIndex(doc, () => false);
