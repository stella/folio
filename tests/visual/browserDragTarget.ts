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

/**
 * Offsets past a paragraph's first and second characters that are not struck,
 * when both are text before any other content that is not struck.
 */
const openingText = (paragraph: {
  forEach: (visit: (child: DragInline, offset: number) => void) => void;
}): readonly [number, number] | null => {
  const after: number[] = [];
  let opensWithText = true;
  paragraph.forEach((child, offset) => {
    if (!opensWithText || after.length >= 2) return;
    if (child.marks.some(({ type }) => type.name === "deletion")) return;
    if (!child.isText) {
      opensWithText = false;
      return;
    }
    for (let index = 1; index <= child.nodeSize && after.length < 2; index += 1)
      after.push(offset + index);
  });
  const [first, second] = after;
  return opensWithText && first !== undefined && second !== undefined
    ? [first, second]
    : null;
};

export const crossParagraphDragRange = (
  doc: DragDocument,
): { from: number; to: number } | null => {
  let previous: number | null = null;
  let range: { from: number; to: number } | null = null;
  doc.forEach((node, offset) => {
    if (range) return;
    const opening = node.type.name === "paragraph" ? openingText(node) : null;
    if (opening !== null && previous !== null)
      range = { from: previous, to: offset + 1 + opening[1] };
    previous = opening === null ? null : offset + 1 + opening[0];
  });
  return range;
};
