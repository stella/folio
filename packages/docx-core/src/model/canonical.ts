import type { Comment, DocumentBody } from "./content";
import type { ParagraphFormatting } from "./formatting";

/** Set modeled comments while keeping an empty comments part out of the document model. */
export const setDocumentBodyComments = (body: DocumentBody, comments: Comment[]): void => {
  if (comments.length === 0) {
    Reflect.set(body, "comments", undefined);
    return;
  }
  body.comments = comments;
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
  return isCanonicallyEmptyParagraphFormatting(formatting) ? undefined : formatting;
};

/** Whether paragraph formatting states no modeled properties. */
export const isCanonicallyEmptyParagraphFormatting = (
  formatting: ParagraphFormatting | undefined,
): boolean =>
  formatting === undefined || Object.keys(formatting).every((key) => key === "sourceXml");
