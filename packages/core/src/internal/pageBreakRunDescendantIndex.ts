import type { Node as PMNode } from "prosemirror-model";

type BuildPageBreakRunDescendantIndexOptions = {
  onNodeVisited?: (node: PMNode) => void;
};

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

/** Build one identity-scoped index for a single layout conversion. */
export const buildPageBreakRunDescendantIndex = (
  doc: PMNode,
  options: BuildPageBreakRunDescendantIndexOptions = {},
): PageBreakRunDescendantIndex => {
  const positions = new WeakMap<PMNode, number>();

  const visitNode = (node: PMNode, startPos: number): number | undefined => {
    options.onNodeVisited?.(node);
    let firstPosition = node.type.name === "pageBreakRun" ? startPos : undefined;

    // Visit every child even after finding the first break. Ancestor answers
    // stay first-in-document-order while each indexed descendant gets its own answer.
    node.forEach((child, offset) => {
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
