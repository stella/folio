export type TextBoundary = { node: Text; offset: number };

const TEXT_NODE_TYPE = 3;

const isTextNode = (node: Node): node is Text => node.nodeType === TEXT_NODE_TYPE;

/** Treat every descendant text node as one logical character stream. */
export const descendantTextNodes = (root: Node): Text[] => {
  const nodes: Text[] = [];
  const visit = (node: Node): void => {
    if (isTextNode(node)) {
      nodes.push(node);
      return;
    }
    for (const child of node.childNodes) visit(child);
  };
  visit(root);
  return nodes;
};

export const totalTextLength = (nodes: readonly Text[]): number => {
  let length = 0;
  for (const node of nodes) length += node.length;
  return length;
};

export const textBoundaryAt = (
  nodes: readonly Text[],
  offset: number,
  edge: "start" | "end",
): TextBoundary | null => {
  let consumed = 0;
  for (const [index, node] of nodes.entries()) {
    const next = consumed + node.length;
    const ownsBoundary =
      offset < next || (edge === "end" && offset === next) || index === nodes.length - 1;
    if (ownsBoundary) {
      return { node, offset: Math.max(0, Math.min(node.length, offset - consumed)) };
    }
    consumed = next;
  }
  return null;
};

export const logicalTextOffset = (
  nodes: readonly Text[],
  target: Node,
  offset: number,
): number | null => {
  let consumed = 0;
  for (const text of nodes) {
    if (text === target) {
      return consumed + Math.max(0, Math.min(text.length, offset));
    }
    consumed += text.length;
  }
  return null;
};

export const createTextStreamRange = (
  element: HTMLElement,
  startOffset: number,
  endOffset: number,
): Range | null => {
  const nodes = descendantTextNodes(element);
  const length = totalTextLength(nodes);
  const boundedStart = Math.max(0, Math.min(length, startOffset));
  const boundedEnd = Math.max(0, Math.min(length, endOffset));
  const start = textBoundaryAt(nodes, boundedStart, "start");
  const end = textBoundaryAt(nodes, boundedEnd, boundedStart === boundedEnd ? "start" : "end");
  if (!start || !end) return null;

  const range = element.ownerDocument.createRange();
  range.setStart(start.node, start.offset);
  range.setEnd(end.node, end.offset);
  return range;
};
