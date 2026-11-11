/** Structural selection targets shared by playground bridges and input traces. */
export type BrowserDragTarget =
  "table" | "list" | "note" | "field" | "inlineObject" | "paragraphs";

type DragInline = {
  isText: boolean;
  nodeSize: number;
  marks: readonly { type: { name: string } }[];
};

/**
 * A selection across two adjacent body paragraphs that each open with text:
 * from after the first's first character to two characters into the next.
 * Struck text is not counted, so the drag names the same text whether the
 * earlier input was applied or suggested.
 */
type DragDocument = {
  forEach: (
    visit: (
      node: {
        type: { name: string };
        forEach: (visit: (child: DragInline, offset: number) => void) => void;
      },
      offset: number,
    ) => void,
  ) => void;
};

/** The offset of a paragraph's first text that is not struck, when it opens with two characters. */
const openingText = (paragraph: {
  forEach: (visit: (child: DragInline, offset: number) => void) => void;
}) => {
  let opening: number | null | undefined;
  paragraph.forEach((child, offset) => {
    if (opening !== undefined) return;
    if (child.marks.some(({ type }) => type.name === "deletion")) return;
    opening = child.isText && child.nodeSize >= 2 ? offset : null;
  });
  return opening ?? null;
};

export const crossParagraphDragRange = (
  doc: DragDocument,
): { from: number; to: number } | null => {
  let previous: number | null = null;
  let range: { from: number; to: number } | null = null;
  doc.forEach((node, offset) => {
    if (range) return;
    const opening = node.type.name === "paragraph" ? openingText(node) : null;
    const start = opening === null ? null : offset + 1 + opening;
    if (start !== null && previous !== null)
      range = { from: previous + 1, to: start + 2 };
    previous = start;
  });
  return range;
};
