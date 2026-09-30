/**
 * Removing a paragraph mark: the one join both a resolved revision and a
 * retracted one perform.
 *
 * A paragraph's mark ends it, and in OOXML its properties live on that mark.
 * With the mark gone, the paragraph runs on into the next one and ends with
 * the NEXT paragraph's mark, so the paragraph left is the next one: its
 * identity, its paragraph properties (style, numbering, the mark's own run
 * properties), its mark's revision and its section endpoint all survive, and
 * the removed mark takes its paragraph's properties and section endpoint with
 * it. This holds whether the first paragraph still has words or not.
 *
 * The first paragraph's runs keep their direct formatting; what its style
 * lent them is re-read in the next paragraph's style.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { joinProseParagraphsWithRightPropertySource } from "../docx/paragraphPropertySource";
import { JOINED_RUNS_RESTYLED_META } from "./extensions/features/JoinedRunStyleExtension";
import { rebaseParagraphRuns } from "./rebaseParagraphRunFormatting";
import { paragraphRunStyleContext, type RunStyleResolver } from "./runStyleFormatting";

type JoinAtParagraphMarkOptions = {
  tr: Transaction;
  /** Position of the paragraph whose mark goes. */
  paragraphPos: number;
  paragraph: PMNode;
  /** The paragraph directly after it, which the join runs into. */
  next: PMNode;
  /**
   * Re-reads the moved runs' inherited formatting in the next paragraph's
   * style. Without one, the editor's join restyler does it when the
   * transaction is dispatched.
   */
  styleResolver: RunStyleResolver | null;
};

export const joinAtParagraphMark = ({
  tr,
  paragraphPos,
  paragraph,
  next,
  styleResolver,
}: JoinAtParagraphMarkOptions): void => {
  joinProseParagraphsWithRightPropertySource({
    attrs: next.attrs,
    pos: paragraphPos + paragraph.nodeSize,
    transaction: tr,
  });
  if (!styleResolver) return;
  if (paragraph.content.size > 0) {
    rebaseParagraphRuns({
      previousContext: paragraphRunStyleContext(paragraph, styleResolver),
      paragraphPosition: paragraphPos,
      range: { from: 0, to: paragraph.content.size },
      styleResolver,
      tr,
    });
  }
  tr.setMeta(JOINED_RUNS_RESTYLED_META, true);
};
