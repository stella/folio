/**
 * The character styles folio's serializers write into every document that
 * carries a comment or a note, and the definitions that keep those references
 * from dangling.
 *
 * `commentSerializer`, `paragraphSerializer` and `noteSerializer` each emit
 * `<w:rStyle w:val="…"/>` for the reference mark, because that is what Word
 * writes and what makes the mark superscript. Nothing guaranteed the document
 * *defined* those styles: a package folio assembled from the generic style set
 * referenced `FootnoteReference` while declaring only seven paragraph styles,
 * so the mark rendered as body text. One owner here, so a serializer and the
 * style table cannot disagree about the id.
 */

import { BUILT_IN_STYLE_NAME } from "./builtInStyles";
import type { Style, StyleDefinitions } from "../types/document";

/** The id `commentSerializer` and `paragraphSerializer` write. */
export const COMMENT_REFERENCE_STYLE_ID = "CommentReference";
/** The ids `noteSerializer` writes for the two note kinds. */
export const FOOTNOTE_REFERENCE_STYLE_ID = "FootnoteReference";
export const ENDNOTE_REFERENCE_STYLE_ID = "EndnoteReference";

/** Which reference marks a package actually contains. */
export type NoteReferenceNeeds = {
  comments: boolean;
  footnotes: boolean;
  endnotes: boolean;
};

/** The default character style Word bases these on, when the package has one. */
const DEFAULT_CHARACTER_STYLE_ID = "DefaultParagraphFont";

const referenceStyle = (styleId: string, name: string, superscript: boolean): Style => ({
  styleId,
  type: "character",
  name,
  uiPriority: 99,
  semiHidden: true,
  unhideWhenUsed: true,
  rPr: superscript ? { vertAlign: "superscript" } : { fontSize: 16 },
});

/**
 * The definition for each reference style, keyed by the need that requires it.
 * Word's comment reference is 8pt body text; the note references are
 * superscript.
 */
const REFERENCE_STYLES: Record<keyof NoteReferenceNeeds, () => Style> = {
  comments: () =>
    referenceStyle(COMMENT_REFERENCE_STYLE_ID, BUILT_IN_STYLE_NAME.commentReference, false),
  footnotes: () =>
    referenceStyle(FOOTNOTE_REFERENCE_STYLE_ID, BUILT_IN_STYLE_NAME.footnoteReference, true),
  endnotes: () =>
    referenceStyle(ENDNOTE_REFERENCE_STYLE_ID, BUILT_IN_STYLE_NAME.endnoteReference, true),
};

/** What a package's content requires, read off the model rather than guessed. */
export const noteReferenceNeeds = (pkg: {
  comments?: unknown[] | undefined;
  footnotes?: unknown[] | undefined;
  endnotes?: unknown[] | undefined;
}): NoteReferenceNeeds => ({
  comments: (pkg.comments?.length ?? 0) > 0,
  footnotes: (pkg.footnotes?.length ?? 0) > 0,
  endnotes: (pkg.endnotes?.length ?? 0) > 0,
});

/**
 * The reference styles a package needs and does not already define, by id.
 *
 * Returns the definitions to append rather than a mutated style table: the
 * caller owns the document, and a document that already defines the style
 * (under any id folio recognises) keeps its own.
 */
export const missingNoteReferenceStyles = (
  styles: StyleDefinitions | undefined,
  needs: NoteReferenceNeeds,
): Style[] => {
  const defined = new Set((styles?.styles ?? []).map((style) => style.styleId));
  // Only chain to the default character style when this package declares one:
  // a `w:basedOn` naming a style that does not exist is the same defect this
  // function exists to remove.
  const basedOn = defined.has(DEFAULT_CHARACTER_STYLE_ID) ? DEFAULT_CHARACTER_STYLE_ID : undefined;
  const missing: Style[] = [];
  for (const [need, create] of Object.entries(REFERENCE_STYLES)) {
    if (!needs[need as keyof NoteReferenceNeeds]) {
      continue;
    }
    const style = create();
    if (!defined.has(style.styleId)) {
      missing.push(basedOn === undefined ? style : { ...style, basedOn });
    }
  }
  return missing;
};
