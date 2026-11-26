import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import {
  getProseParagraphPropertySourceToken,
  joinProseParagraphsWithRightPropertySource,
  markParagraphPropertySourceTransfers,
} from "../../docx/paragraphPropertySource";
import { paragraphJoinAfter } from "../../prosemirror/commands/paragraphBookmarkJoin";
import { JOINED_RUNS_RESTYLED_META } from "../../prosemirror/extensions/features/JoinedRunStyleExtension";
import { joinedParagraphAttrs } from "../../prosemirror/paragraphMarkJoin";
import {
  paragraphAt,
  rebaseParagraphRuns,
} from "../../prosemirror/rebaseParagraphRunFormatting";
import {
  paragraphRunStyleContext,
  type RunStyleResolver,
} from "../../prosemirror/runStyleFormatting";

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
  const boundary = paragraphPos + paragraph.nodeSize;
  const join = paragraphJoinAfter(tr.doc, boundary);
  if (!join || join.next !== next)
    panic("A paragraph-mark join lost its next paragraph");
  if (join.inlineBookmarks.length > 0) {
    const leftToken = getProseParagraphPropertySourceToken(paragraph);
    const rightToken = getProseParagraphPropertySourceToken(next);
    tr.replaceWith(boundary - 1, join.nextPosition + 1, join.inlineBookmarks);
    tr.setNodeMarkup(
      paragraphPos,
      undefined,
      joinedParagraphAttrs({
        paragraph,
        next,
        between: join.inlineBookmarks.reduce(
          (size, node) => size + node.nodeSize,
          0,
        ),
        styleResolver,
      }),
    );
    markParagraphPropertySourceTransfers(tr, [
      {
        displacedToken: typeof leftToken === "string" ? leftToken : null,
        selectedToken: typeof rightToken === "string" ? rightToken : null,
      },
    ]);
  } else {
    joinProseParagraphsWithRightPropertySource({
      attrs: joinedParagraphAttrs({ paragraph, next, styleResolver }),
      pos: boundary,
      transaction: tr,
    });
  }
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

type RebaseParagraphRunFormattingOptions = {
  nextAttrs: Record<string, unknown>;
  paragraphPosition: number;
  shouldRebase?: (node: PMNode) => boolean;
  styleResolver: RunStyleResolver;
  tr: Transaction;
};

/**
 * Change paragraph attrs and re-resolve inherited run marks in the new style
 * context without turning the old rendered style into direct run formatting.
 */
export const setParagraphAttrsWithRebasedRunFormatting = ({
  nextAttrs,
  paragraphPosition,
  shouldRebase,
  styleResolver,
  tr,
}: RebaseParagraphRunFormattingOptions): Transaction => {
  const previousContext = paragraphRunStyleContext(
    paragraphAt(tr, paragraphPosition),
    styleResolver,
  );
  tr = tr.setNodeMarkup(paragraphPosition, undefined, nextAttrs);
  return rebaseParagraphRuns({
    previousContext,
    paragraphPosition,
    ...(shouldRebase ? { shouldRebase } : {}),
    styleResolver,
    tr,
  });
};

/** The reference allocator uses this seed when a transaction needs reproducible IDs. */
export const deterministicParaIdSeedKey =
  "stll.folio.meta.deterministicParaIdSeed";

/** Ask the allocator to mint reproducible ids for the paragraphs `tr` creates. */
export const requestDeterministicParaIds = (
  tr: Transaction,
  seed: string,
): Transaction => tr.setMeta(deterministicParaIdSeedKey, seed);
