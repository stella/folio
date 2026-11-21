/**
 * Comment paragraphs that repeat a paragraph id, given ids of their own.
 *
 * `comments.xml` is not a story part, so `ensureParaIds` leaves it as the
 * package holds it. A producer that copies a comment paragraph copies its
 * `w14:paraId` too, and two paragraphs answering to one id are refused by the
 * canonical session, which addresses paragraphs by it. The parser therefore
 * gives every repeat after the first a minted id, keyed as `ensureParaIds`
 * keys one (the part, the paragraph's markup, how many earlier paragraphs of
 * the part repeat that markup), so both engines read the same ids from the
 * same bytes. Thread links were resolved from the ids as written before this
 * runs, and a save writes the minted id into both comment parts together.
 */
import { PARSE_WARNING_CODES } from "@stll/docx-core/model";
import { parseStreamingXmlWithSourceRanges } from "@stll/docx-core/xml";

import type { Comment, Paragraph } from "../types/document";
import {
  isWordprocessingElement,
  type MintContext,
  mintParaId,
  paragraphMintKeys,
  storyParagraphSlots,
} from "./ensureParaIds";
import {
  getAttributeByNamespaceUri,
  getChildElements,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

/** The code this repair is reported under, owned here, not at the caller. */
export const REPEATED_COMMENT_PARAGRAPH_ID_WARNING: "repeated-comment-paragraph-id" =
  PARSE_WARNING_CODES.repeatedCommentParagraphId;

const DECIMAL_NUMBER = /^[+-]?\d+$/u;

/**
 * Each comment's direct paragraphs as their ordinals among the part's
 * paragraph slots, by the `w:id` of the first comment defining it: the
 * comment the parser keeps.
 */
const commentParagraphOrdinals = (
  xml: string,
): ReadonlyMap<number, readonly number[]> | undefined => {
  const parsed = parseStreamingXmlWithSourceRanges(xml);
  if (parsed.status !== "parsed") return undefined;
  const ordinals = new Map<XmlElement, number>();
  for (const [index, slot] of storyParagraphSlots(parsed.value).entries()) {
    if (slot.type === "paragraph") ordinals.set(slot.element, index + 1);
  }
  const byComment = new Map<number, number[]>();
  for (const comment of getChildElements(parsed.value)) {
    if (!isWordprocessingElement(comment, "comment")) continue;
    const rawId = getAttributeByNamespaceUri(
      comment,
      WORDPROCESSINGML_NAMESPACE_URIS,
      "id",
    );
    if (rawId === null || !DECIMAL_NUMBER.test(rawId)) continue;
    const id = Number.parseInt(rawId, 10);
    if (byComment.has(id)) continue;
    byComment.set(
      id,
      getChildElements(comment).flatMap((child) => {
        const ordinal = ordinals.get(child);
        return ordinal === undefined ? [] : [ordinal];
      }),
    );
  }
  return byComment;
};

type RepeatedCommentParagraph = {
  comment: Comment;
  /** Its index in the comment's content. */
  index: number;
  paragraph: Paragraph;
  /** Its place among every comment's paragraphs, from 1. */
  position: number;
};

type RemintRepeatedCommentParagraphIdsOptions = {
  comments: readonly Comment[];
  /** `word/comments.xml` as the package holds it. */
  commentsXml: string;
  partPath: string;
  /** Every paragraph id the package holds or the parse assigned; minted ids join it. */
  taken: Set<string>;
};

/**
 * Give each comment paragraph whose id an earlier comment paragraph holds a
 * minted one, the first holder keeping it. Answers how many were given one.
 */
export const remintRepeatedCommentParagraphIds = ({
  comments,
  commentsXml,
  partPath,
  taken,
}: RemintRepeatedCommentParagraphIdsOptions): number => {
  const seen = new Set<string>();
  const repeats: RepeatedCommentParagraph[] = [];
  let position = 0;
  for (const comment of comments) {
    for (const [index, paragraph] of comment.content.entries()) {
      position += 1;
      const id = paragraph.paraId?.toUpperCase();
      if (id === undefined) continue;
      if (seen.has(id)) repeats.push({ comment, index, paragraph, position });
      else seen.add(id);
    }
  }
  if (repeats.length === 0) return 0;

  const ordinals = commentParagraphOrdinals(commentsXml);
  const keys = paragraphMintKeys(commentsXml);
  const context: MintContext = { taken };
  for (const { comment, index, paragraph, position: at } of repeats) {
    const own = ordinals?.get(comment.id);
    const ordinal =
      own?.length === comment.content.length ? own[index] : undefined;
    // Without the part's paragraphs to align with, a repeat is keyed on its
    // place among the comments' paragraphs, which the same bytes always give.
    const id =
      ordinal === undefined
        ? mintParaId(context, partPath, at, undefined)
        : mintParaId(context, partPath, ordinal, keys);
    paragraph.paraId = id;
    paragraph.textId = id;
  }
  return repeats.length;
};
