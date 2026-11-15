import type { Comment, DocumentBody } from "./content";
import type { ParagraphFormatting, SpacingExplicit } from "./formatting";

/** Set modeled comments while keeping an empty comments part out of the document model. */
export const setDocumentBodyComments = (
  body: DocumentBody,
  comments: Comment[],
): void => {
  if (comments.length === 0) {
    Reflect.set(body, "comments", undefined);
    return;
  }
  body.comments = comments;
};

/**
 * The sides of `w:spacing` a paragraph's own properties state, as the parser
 * reads them: a side is explicit exactly when the record holds its value.
 *
 * The one rule for the field. The parser marks a record it reads with it, and
 * an operation that states a paragraph's spacing states the field with it, so
 * a paragraph an edit changed has the field the saved file reopens to.
 */
export const paragraphSpacingExplicit = (
  formatting: ParagraphFormatting,
): SpacingExplicit | undefined => {
  const before = formatting.spaceBefore !== undefined;
  const after = formatting.spaceAfter !== undefined;
  if (!before && !after) return undefined;
  return {
    ...(before ? { before: true } : {}),
    ...(after ? { after: true } : {}),
  };
};

/**
 * Finish paragraph formatting at a model boundary.
 *
 * The parser represents an empty `w:pPr`, including one whose only state was
 * source-replay metadata, as absent formatting. Live conversion must use the
 * same representation so save/reopen does not change the model's shape.
 */
export const finalizeParagraphFormatting = (
  formatting: ParagraphFormatting | undefined,
): ParagraphFormatting | undefined => {
  return isCanonicallyEmptyParagraphFormatting(formatting)
    ? undefined
    : formatting;
};

/** Whether paragraph formatting states no modeled properties. */
export const isCanonicallyEmptyParagraphFormatting = (
  formatting: ParagraphFormatting | undefined,
): boolean =>
  formatting === undefined ||
  Object.keys(formatting).every((key) => key === "sourceXml");
