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
 * every record an edit produces is finished with it, so a paragraph an edit
 * changed has the field the saved file reopens to.
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

const statesSameSpacingSides = (
  left: SpacingExplicit | undefined,
  right: SpacingExplicit | undefined,
): boolean =>
  (left?.before === true) === (right?.before === true) &&
  (left?.after === true) === (right?.after === true) &&
  (left === undefined) === (right === undefined);

/**
 * Finish paragraph formatting at a model boundary.
 *
 * The parser represents an empty `w:pPr`, including one whose only state was
 * source-replay metadata, as absent formatting, marks the spacing sides the
 * record states, and reads a hanging indent as a negative first line. Live
 * conversion and every edit must use the same representation so save/reopen
 * does not change the model's shape.
 */
export const finalizeParagraphFormatting = (
  formatting: ParagraphFormatting | undefined,
): ParagraphFormatting | undefined => {
  if (formatting === undefined) return undefined;
  let finished = formatting;
  const edit = (): ParagraphFormatting => {
    if (finished === formatting) finished = { ...formatting };
    return finished;
  };
  const spacingExplicit = paragraphSpacingExplicit(formatting);
  if (!statesSameSpacingSides(spacingExplicit, formatting.spacingExplicit)) {
    if (spacingExplicit === undefined) delete edit().spacingExplicit;
    else edit().spacingExplicit = spacingExplicit;
  }
  // `w:ind` writes a first line or a hanging magnitude, and the parser reads a
  // hanging one back as a first line pulled out to the left, flagged hanging.
  // A flag without a first line writes nothing, and an unset flag is how the
  // parser reads a first line.
  const { indentFirstLine, hangingIndent } = formatting;
  if (hangingIndent !== undefined && hangingIndent !== true)
    delete edit().hangingIndent;
  if (hangingIndent === true) {
    if (indentFirstLine === undefined) delete edit().hangingIndent;
    else {
      const asRead = indentFirstLine === 0 ? 0 : -Math.abs(indentFirstLine);
      if (asRead !== indentFirstLine) edit().indentFirstLine = asRead;
    }
  }
  return isCanonicallyEmptyParagraphFormatting(finished) ? undefined : finished;
};

/** Whether paragraph formatting states no modeled properties. */
export const isCanonicallyEmptyParagraphFormatting = (
  formatting: ParagraphFormatting | undefined,
): boolean =>
  formatting === undefined ||
  Object.keys(formatting).every((key) => key === "sourceXml");
