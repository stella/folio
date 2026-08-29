/**
 * Pure ProseMirror text-extraction helpers shared by the React and Vue
 * adapters.
 *
 * "Vanilla view" text extraction skips text inside `insertion` marks
 * (tracked-change additions that aren't accepted yet) so the agent's view
 * of the document matches what `addComment` / `proposeChange` can anchor
 * to. Tracked deletions stay included — they're still in the doc until
 * accepted.
 *
 * Paragraph range lookup by `w14:paraId` lives in
 * `./utils/findParagraphByParaId` (`findParagraphByParaId`); this file
 * carries only the text-slicing helpers.
 */

import type { Node as PMNode } from "prosemirror-model";

/** Text of a single PM node (typically a paragraph), vanilla view. */
export function getVanillaNodeText(node: PMNode): string {
  const parts: string[] = [];
  node.descendants((child) => {
    if (child.marks.some((m) => m.type.name === "insertion")) return false;
    if (child.isText && child.text) {
      parts.push(child.text);
      return false;
    }
    if (child.isLeaf && child.textContent) {
      parts.push(child.textContent);
      return false;
    }
    return true;
  });
  return parts.join("");
}

/** Text between two doc positions, vanilla view. */
export function getVanillaTextBetween(doc: PMNode, from: number, to: number): string {
  if (from >= to) return "";
  const parts: string[] = [];
  doc.nodesBetween(from, to, (child, pos) => {
    if (child.marks.some((m) => m.type.name === "insertion")) return false;
    if (child.isText && child.text) {
      const start = Math.max(from, pos);
      const end = Math.min(to, pos + child.text.length);
      if (start < end) parts.push(child.text.slice(start - pos, end - pos));
      return false;
    }
    if (child.isLeaf && child.textContent) {
      parts.push(child.textContent);
      return false;
    }
    return true;
  });
  return parts.join("");
}

type TextPosition = {
  text: string;
  pos: number;
  pmLength: number;
  atomic: boolean;
};

/**
 * Find `searchText` within a PM paragraph range and return its position.
 *
 * Returns null if:
 *   - searchText is empty
 *   - searchText is not found
 *   - searchText appears more than once (ambiguous; caller disambiguates)
 *
 * The fullText is built from PM text nodes only and matches the vanilla
 * view the agent reads via `read_document`: tracked insertions are
 * excluded (not in the doc yet), tracked deletions are included (still
 * in the doc until accepted), and comment markers are stripped.
 */
export function findTextInPmParagraph(
  doc: PMNode,
  paragraphFrom: number,
  paragraphTo: number,
  searchText: string,
): { from: number; to: number } | null {
  if (!searchText) return null;

  let fullText = "";
  const textPositions: TextPosition[] = [];

  doc.nodesBetween(paragraphFrom, paragraphTo, (node, pos) => {
    if (node.marks.some((m) => m.type.name === "insertion")) return false;
    if (node.isText && node.text) {
      textPositions.push({ text: node.text, pos, pmLength: node.text.length, atomic: false });
      fullText += node.text;
      return false;
    }
    if (node.isLeaf && node.textContent) {
      textPositions.push({ text: node.textContent, pos, pmLength: node.nodeSize, atomic: true });
      fullText += node.textContent;
      return false;
    }
    return true;
  });

  const firstMatch = fullText.indexOf(searchText);
  if (firstMatch === -1) return null;
  // Reject ambiguous searches — the LLM gets a clearer error than a silent mistarget.
  const secondMatch = fullText.indexOf(searchText, firstMatch + 1);
  if (secondMatch !== -1) return null;

  const matchEnd = firstMatch + searchText.length;
  let segmentStart = 0;
  for (const position of textPositions) {
    const segmentEnd = segmentStart + position.text.length;
    if (
      position.atomic &&
      ((segmentStart < firstMatch && firstMatch < segmentEnd) ||
        (segmentStart < matchEnd && matchEnd < segmentEnd))
    ) {
      return null;
    }
    segmentStart = segmentEnd;
  }

  // Map string offset back to PM position.
  let charOffset = 0;
  let fromPos = paragraphFrom;
  let toPos = paragraphFrom;

  for (const tp of textPositions) {
    const segEnd = charOffset + tp.text.length;
    if (charOffset <= firstMatch && firstMatch < segEnd) {
      const localOffset = firstMatch - charOffset;
      fromPos = tp.atomic ? tp.pos : tp.pos + localOffset;
    }
    if (charOffset <= firstMatch + searchText.length && firstMatch + searchText.length <= segEnd) {
      const localOffset = firstMatch + searchText.length - charOffset;
      toPos = tp.atomic ? tp.pos + tp.pmLength : tp.pos + localOffset;
      break;
    }
    charOffset = segEnd;
  }

  return { from: fromPos, to: toPos };
}
