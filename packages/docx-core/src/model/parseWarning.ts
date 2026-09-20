/**
 * What folio normalised to open a document, as data rather than as prose.
 *
 * Word accepts input the schema does not describe, so folio accepts it too and
 * normalises at the parse boundary. Every such decision is recorded here: a
 * stable `code`, the part it happened in, the best position that part can
 * name, and the value folio declined to read. A host can count them, group
 * them, or decide on one; the human-readable string the parser has always
 * exposed is derived from this, never written by hand at the call site, so the
 * two cannot drift.
 *
 * It lives beside `Document` because it is part of the parsed document, and a
 * `Document` may be produced by anything that builds this model.
 */

/**
 * Every normalisation the parse boundary is allowed to make.
 *
 * Adding a member without also giving it a message is a type error: the
 * formatter map over this union is total.
 */
export const PARSE_WARNING_CODES = {
  /** A package opened from encrypted storage; a save writes it unencrypted. */
  packageDecrypted: "package-decrypted",
  /** The archive reader reported something about the package it opened. */
  packageArchive: "package-archive",
  /** The main document part is missing; the body is empty. */
  documentPartMissing: "document-part-missing",
  /** The model validator reported an issue the parse did not treat as fatal. */
  documentModelIssue: "document-model-issue",
  /** A later `w:comment` repeating a `w:id` an earlier one defines. */
  duplicateCommentId: "duplicate-comment-id",
  /** A `w:comment` with no readable `w:id`; it can anchor nothing. */
  missingCommentId: "missing-comment-id",
  /** A later `w:footnote`/`w:endnote` repeating an id an earlier one defines. */
  duplicateNoteId: "duplicate-note-id",
  /** A comment marker naming a comment `comments.xml` never defines. */
  danglingCommentReference: "dangling-comment-reference",
  /** A comment range with no matching start or end, re-anchored as a point. */
  unbalancedCommentRange: "unbalanced-comment-range",
  /** A `w:headerReference` naming a header part the package does not carry. */
  danglingHeaderReference: "dangling-header-reference",
  /** A `w:footerReference` naming a footer part the package does not carry. */
  danglingFooterReference: "dangling-footer-reference",
  /** A `r:id` naming a relationship the part's `.rels` does not define. */
  danglingRelationshipId: "dangling-relationship-id",
  /** A paragraph naming a `w:num` the numbering part never defines. */
  unnumberedParagraph: "unnumbered-paragraph",
  /** A style naming a `w:num` the numbering part never defines. */
  unnumberedStyle: "unnumbered-style",
  /** A tracked move range with no matching start or end. */
  unbalancedMoveRange: "unbalanced-move-range",
  /** A `w:type` outside `ST_HdrFtr`, read as the enumeration's default. */
  headerFooterTypeOutsideEnum: "header-footer-type-outside-enum",
  /** A value outside `ST_OnOff`'s six spellings. */
  unrecognisedOnOffValue: "unrecognised-on-off-value",
  /** A theme-colour token outside `ST_ThemeColor`, kept but not painted. */
  unrecognisedThemeColor: "unrecognised-theme-color",
  /** A `CT_Border` with no `w:val`, which has no border style to read. */
  borderWithoutValue: "border-without-value",
  /** A style set defining two styles under one id. */
  styleSetDuplicateStyleId: "style-set-duplicate-style-id",
  /** A style set naming an initial paragraph style it does not contain. */
  styleSetInitialStyleMissing: "style-set-initial-style-missing",
  /** An explicit page-break run the editable model lays out approximately. */
  pageBreakProjectionApproximated: "page-break-projection-approximated",
} as const;

export type ParseWarningCode = (typeof PARSE_WARNING_CODES)[keyof typeof PARSE_WARNING_CODES];

/** Where a warning happened, as precisely as the part can say. */
export type ParseWarningLocation = {
  /** The package part, e.g. `word/document.xml`. `package` for the container. */
  part: string;
  /** The element as written, prefix included, when one is known. */
  element?: string;
  /**
   * The best position the part can name: `paragraph 12`, `style "Heading1"`,
   * `w:id 4`. Free-form because what identifies a place differs per part, and
   * a byte offset is not available once the tree is built.
   */
  at?: string;
};

export type ParseWarning = {
  code: ParseWarningCode;
  location: ParseWarningLocation;
  /** The value folio declined to read, exactly as the source wrote it. */
  value?: string;
  /** How many occurrences this entry stands for; at least 1. */
  count: number;
  /**
   * Text from the owner that produced it, for the codes that pass a message
   * through from the archive reader or the model validator.
   */
  detail?: string;
};

/**
 * The cap on retained warnings per code.
 *
 * A generated document can repeat one defect on every paragraph, and a host
 * that only wants to know a defect occurred should not pay for a hundred
 * thousand records. Occurrences past the cap are counted, not kept.
 */
export const MAX_RETAINED_PARSE_WARNINGS_PER_CODE = 20;
