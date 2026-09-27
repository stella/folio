/**
 * Removing a paragraph mark: the one join both a resolved revision and a
 * retracted one perform.
 *
 * A paragraph's mark ends it, and in OOXML its properties live on that mark.
 * With the mark gone, the paragraph runs on into the next one and ends with
 * the NEXT paragraph's mark: that mark's own revision and its section
 * endpoint survive, and the removed mark's section endpoint goes with it.
 * The joined paragraph keeps the first paragraph's properties unless the
 * first held nothing, in which case only the join survives of it and the
 * next paragraph is what the reader is left with.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { joinProseParagraphsWithRightPropertySource } from "../docx/paragraphPropertySource";
import { holdsNoContent } from "./zeroWidthAnchors";

type JoinAtParagraphMarkOptions = {
  tr: Transaction;
  /** Position of the paragraph whose mark goes. */
  paragraphPos: number;
  paragraph: PMNode;
  /** The paragraph directly after it, which the join runs into. */
  next: PMNode;
  /**
   * The first paragraph's content is going too (a deleted paragraph), so
   * only the join survives of it, as when it is already empty.
   */
  firstIsGoing?: boolean;
};

export const joinAtParagraphMark = ({
  tr,
  paragraphPos,
  paragraph,
  next,
  firstIsGoing = false,
}: JoinAtParagraphMarkOptions): void => {
  const joinPos = paragraphPos + paragraph.nodeSize;
  const emptyFirstParagraph = firstIsGoing || holdsNoContent(paragraph);
  const formattingOwner = emptyFirstParagraph ? next : paragraph;
  const joinedAttrs = {
    ...formattingOwner.attrs,
    pPrMark: next.attrs["pPrMark"],
    _sectionProperties: next.attrs["_sectionProperties"],
  };
  if (emptyFirstParagraph) {
    joinProseParagraphsWithRightPropertySource({
      attrs: joinedAttrs,
      pos: joinPos,
      transaction: tr,
    });
    return;
  }
  tr.join(joinPos);
  tr.setNodeMarkup(paragraphPos, undefined, joinedAttrs);
};
