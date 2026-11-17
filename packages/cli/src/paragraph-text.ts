/**
 * Text edits inside one paragraph's XML that ignore how its text is split
 * into runs: a match may start in one `w:t` and end in another, and the
 * replacement takes the formatting of the run the match starts in. Deleted
 * and moved-away text and field instructions are not part of the text.
 */

import { childElements, isW, type XmlElement } from "./ooxml-package";

/** Elements whose text is not the paragraph's visible text. */
const HIDDEN_TEXT_CONTAINERS = ["del", "moveFrom", "instrText", "delText"];

type TextNode = { element: XmlElement; text: string };

const textOf = (element: XmlElement): string =>
  (element.elements ?? [])
    .map((node) => (node.type === "text" ? String(node.text ?? "") : ""))
    .join("");

const textNodes = (paragraph: XmlElement): TextNode[] => {
  const nodes: TextNode[] = [];
  const visit = (element: XmlElement): void => {
    for (const child of childElements(element)) {
      if (HIDDEN_TEXT_CONTAINERS.some((name) => isW(child, name))) continue;
      // A nested paragraph (a text box's) is its own paragraph.
      if (isW(child, "p") || isW(child, "txbxContent")) continue;
      if (isW(child, "t")) nodes.push({ element: child, text: textOf(child) });
      else visit(child);
    }
  };
  visit(paragraph);
  return nodes;
};

const setText = (element: XmlElement, text: string): void => {
  element.elements = text === "" ? [] : [{ type: "text", text }];
  element.attributes = { ...element.attributes, "xml:space": "preserve" };
};

/** The visible text of a paragraph, as its `w:t` elements spell it. */
export const paragraphText = (paragraph: XmlElement): string =>
  textNodes(paragraph)
    .map(({ text }) => text)
    .join("");

/** One replacement: offsets into {@link paragraphText}. */
export type TextReplacement = { start: number; end: number; text: string };

/**
 * Apply replacements to a paragraph's text. Overlapping replacements after
 * the first are ignored. Returns how many were applied.
 */
export const replaceParagraphText = (
  paragraph: XmlElement,
  replacements: readonly TextReplacement[],
): number => {
  const nodes = textNodes(paragraph);
  const starts: number[] = [];
  let offset = 0;
  for (const { text } of nodes) {
    starts.push(offset);
    offset += text.length;
  }
  const ordered: TextReplacement[] = [];
  for (const replacement of replacements.toSorted(
    (left, right) => left.start - right.start,
  )) {
    const previous = ordered.at(-1);
    if (
      replacement.end > replacement.start &&
      replacement.end <= offset &&
      (previous === undefined || replacement.start >= previous.end)
    )
      ordered.push(replacement);
  }
  const nodeAt = (position: number): number =>
    starts.findLastIndex(
      (start, index) =>
        start <= position &&
        position < start + (nodes[index]?.text.length ?? 0),
    );
  for (const { start, end, text } of ordered.toReversed()) {
    const first = nodeAt(start);
    const last = nodeAt(end - 1);
    const firstNode = nodes[first];
    const lastNode = nodes[last];
    if (firstNode === undefined || lastNode === undefined) continue;
    const firstStart = starts[first] ?? 0;
    const lastStart = starts[last] ?? 0;
    if (first === last) {
      firstNode.text =
        firstNode.text.slice(0, start - firstStart) +
        text +
        firstNode.text.slice(end - firstStart);
      setText(firstNode.element, firstNode.text);
      continue;
    }
    lastNode.text = lastNode.text.slice(end - lastStart);
    setText(lastNode.element, lastNode.text);
    for (let index = first + 1; index < last; index += 1) {
      const middle = nodes[index];
      if (middle === undefined) continue;
      middle.text = "";
      setText(middle.element, "");
    }
    firstNode.text = firstNode.text.slice(0, start - firstStart) + text;
    setText(firstNode.element, firstNode.text);
  }
  return ordered.length;
};
