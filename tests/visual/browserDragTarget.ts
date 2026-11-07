/** Structural selection targets shared by playground bridges and input traces. */
export type BrowserDragTarget =
  "table" | "list" | "note" | "field" | "inlineObject" | "paragraphs";

/**
 * A selection across two adjacent body paragraphs that each open with text:
 * from after the first's first character to two characters into the next.
 */
type DragDocument = {
  forEach: (
    visit: (
      node: {
        type: { name: string };
        firstChild: { isText: boolean; nodeSize: number } | null;
      },
      offset: number,
    ) => void,
  ) => void;
};

export const crossParagraphDragRange = (
  doc: DragDocument,
): { from: number; to: number } | null => {
  let previous: number | null = null;
  let range: { from: number; to: number } | null = null;
  doc.forEach((node, offset) => {
    if (range) return;
    const first = node.firstChild;
    const opensWithText =
      node.type.name === "paragraph" &&
      first?.isText === true &&
      first.nodeSize >= 2;
    if (opensWithText && previous !== null)
      range = { from: previous + 2, to: offset + 3 };
    previous = opensWithText ? offset : null;
  });
  return range;
};
