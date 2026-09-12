import type { Mark } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { expectHyperlinkMarkAttrs } from "./attrs";

type HyperlinkOccurrence = {
  readonly mark: Mark;
  readonly from: number;
  to: number;
};

/**
 * Canonicalize the private identity that keeps adjacent equal hyperlinks apart.
 *
 * Revision resolution can remove a hyperlink before another one. Its ordinal
 * must then close the gap so the live tree equals the next model projection;
 * semantic link identity remains in the hyperlink's public attrs. Package
 * relationship ids are transport details and are deliberately excluded from
 * this identity.
 */
export const canonicalizeHyperlinkOccurrenceIndexes = (tr: Transaction): Transaction => {
  const occurrences: HyperlinkOccurrence[] = [];
  tr.doc.descendants((node, position) => {
    if (!node.isInline) return;
    const hyperlink = node.marks.find(({ type }) => type.name === "hyperlink");
    if (!hyperlink) return;
    const previous = occurrences.at(-1);
    if (previous && previous.to === position && previous.mark.eq(hyperlink)) {
      previous.to += node.nodeSize;
      return;
    }
    occurrences.push({ mark: hyperlink, from: position, to: position + node.nodeSize });
  });
  for (const [occurrence, range] of occurrences.entries()) {
    if (expectHyperlinkMarkAttrs(range.mark)._docxHyperlinkIndex === occurrence) continue;
    tr.removeMark(range.from, range.to, range.mark.type);
    tr.addMark(
      range.from,
      range.to,
      range.mark.type.create({ ...range.mark.attrs, _docxHyperlinkIndex: occurrence }),
    );
  }
  return tr;
};
