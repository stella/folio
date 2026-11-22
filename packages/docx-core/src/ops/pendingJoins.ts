/**
 * Paragraphs a pending mark deletion joins. A paragraph whose mark is deleted
 * or moved away reads as one paragraph with the paragraph after it, which
 * accepting makes it, so the words of the paragraphs joined so far are the
 * words a further join sees.
 */

import type {
  BlockContent,
  Paragraph,
  ParagraphFormatting,
} from "../model/document";
import { structurallyEqual } from "./equality";
import { leafSpans } from "./leaves";
import { paragraphLength, paragraphLogicalText } from "./offsets";
import { isRemovedRevisionNode, paragraphPropertiesOf } from "./review";
import {
  DOCUMENT_OP_TYPES,
  EMPTY_PROPERTY_SETS,
  type OpStory,
  type SetParagraphPropsOp,
} from "./types";

/** Accepted-view text without changing the canonical paragraph or its offset space. */
export const paragraphVisibleText = (paragraph: Paragraph): string => {
  const text = paragraphLogicalText(paragraph);
  return leafSpans(paragraph.content)
    .filter(({ ancestors }) => !ancestors.some(isRemovedRevisionNode))
    .map(({ before, after }) => text.slice(before.offset, after.offset))
    .join("");
};

/** Whether a paragraph's mark is pending deletion, as a move's source too. */
export const joinsFollowingParagraph = ({ pPrMark }: Paragraph): boolean =>
  pPrMark?.kind === "del" || pPrMark?.kind === "moveFrom";

/**
 * Whether the paragraphs right before `blocks[index]`, each joined to the
 * next by a pending mark deletion, run words into it.
 */
export const pendingJoinsRunWordsInto = (
  blocks: readonly BlockContent[],
  index: number,
): boolean => {
  for (let before = index - 1; before >= 0; before -= 1) {
    const block = blocks[before];
    if (block?.type !== "paragraph" || !joinsFollowingParagraph(block))
      return false;
    if (paragraphVisibleText(block) !== "") return true;
  }
  return false;
};

type ParagraphPropertiesSettingOptions = {
  story: OpStory;
  blockId: string;
  /** The properties the paragraph states now. */
  from: ParagraphFormatting | undefined;
  /** The properties it takes; its mark's run properties stay its own. */
  to: ParagraphFormatting | undefined;
};

/** The operation giving a paragraph another's paragraph properties; none when they are the same. */
export const paragraphPropertiesSetting = ({
  story,
  blockId,
  from,
  to,
}: ParagraphPropertiesSettingOptions): SetParagraphPropsOp | undefined => {
  const current = paragraphPropertiesOf(from);
  const desired = paragraphPropertiesOf(to);
  if (structurallyEqual(current, desired)) return undefined;
  return {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story,
    blockId,
    patch: Object.fromEntries([
      ...Object.keys(current ?? {}).map((key) => [key, null]),
      ...Object.entries(desired ?? {}),
    ]),
    whenEmpty:
      to === undefined ? EMPTY_PROPERTY_SETS.OMIT : EMPTY_PROPERTY_SETS.KEEP,
  };
};

type JoinedPropertiesSourceOptions = {
  /** The block list, its selected content deleted and no paragraph joined yet. */
  blocks: readonly BlockContent[];
  first: number;
  last: number;
};

/**
 * The paragraph whose properties the paragraphs from `first` to `last` have
 * once joined, where a pending mark deletion already runs words into them or
 * joins two of them: the one ending the first paragraph the editing view
 * shows there with words, else the last. `undefined` where joining them one
 * by one gives those properties: no such mark, or one on the last paragraph,
 * which takes the properties of the paragraph after it.
 */
export const joinedPropertiesSource = ({
  blocks,
  first,
  last,
}: JoinedPropertiesSourceOptions): Paragraph | undefined => {
  let source: Paragraph | undefined;
  let pending = pendingJoinsRunWordsInto(blocks, first);
  let holdsWords = pending;
  for (let index = first; index <= last; index += 1) {
    const paragraph = blocks[index];
    if (paragraph?.type !== "paragraph") return undefined;
    const joins = joinsFollowingParagraph(paragraph);
    if (index === last) {
      if (joins || !pending) return undefined;
      return source ?? paragraph;
    }
    pending ||= joins;
    holdsWords ||= paragraphLength(paragraph) > 0;
    if (joins) continue;
    if (holdsWords) source ??= paragraph;
    holdsWords = false;
  }
  return undefined;
};
