/**
 * The one place a parse warning becomes a sentence.
 *
 * `Document.warnings` is the string list hosts have always read, and it is now
 * rendered from `Document.parseWarnings` rather than written at the call site,
 * so a warning cannot say one thing in the structured list and another in the
 * prose. The map is total over the code union: a new code does not compile
 * until it has a message.
 */

import {
  PARSE_WARNING_CODES,
  type ParseWarning,
  type ParseWarningCode,
} from "@stll/docx-core/model";

const plural = (count: number, singular: string, pluralForm = `${singular}s`): string =>
  `${String(count)} ${count === 1 ? singular : pluralForm}`;

/** The location suffix, omitted when the part is all we know. */
const where = ({ location }: ParseWarning): string => {
  if (location.at === undefined) {
    return "";
  }
  return ` at ${location.at} in ${location.part}`;
};

const quoted = (value: string | undefined): string => (value === undefined ? "" : ` "${value}"`);

type ParseWarningFormatter = (warning: ParseWarning) => string;

const PARSE_WARNING_MESSAGES = {
  [PARSE_WARNING_CODES.packageDecrypted]: () =>
    "Document was opened from password-protected storage; saving writes an unencrypted .docx file.",
  [PARSE_WARNING_CODES.packageArchive]: (warning) =>
    warning.detail ?? "The archive reader reported an issue.",
  [PARSE_WARNING_CODES.documentPartMissing]: () => "No document.xml found in DOCX",
  [PARSE_WARNING_CODES.documentModelIssue]: (warning) =>
    warning.detail ?? "The document model reported an issue.",
  [PARSE_WARNING_CODES.duplicateCommentId]: (warning) =>
    `Dropped ${plural(warning.count, "comment")} repeating a w:id another comment already defines.`,
  [PARSE_WARNING_CODES.missingCommentId]: (warning) =>
    `Dropped ${plural(warning.count, "comment")} with no readable w:id.`,
  [PARSE_WARNING_CODES.duplicateNoteId]: (warning) =>
    `Dropped ${plural(warning.count, "note")} repeating a w:id another note already defines.`,
  [PARSE_WARNING_CODES.danglingCommentReference]: (warning) =>
    `Removed ${plural(warning.count, "dangling comment reference marker")} whose comments.xml entries are missing.`,
  [PARSE_WARNING_CODES.unbalancedCommentRange]: (warning) =>
    `Re-anchored ${plural(warning.count, "unbalanced comment range marker")} as point comments.`,
  [PARSE_WARNING_CODES.danglingHeaderReference]: (warning) =>
    `Removed ${plural(warning.count, "dangling header reference")} whose header parts are missing.`,
  [PARSE_WARNING_CODES.danglingFooterReference]: (warning) =>
    `Removed ${plural(warning.count, "dangling footer reference")} whose footer parts are missing.`,
  [PARSE_WARNING_CODES.danglingRelationshipId]: (warning) =>
    `Left relationship id${quoted(warning.value)} unresolved${where(warning)}; the part defines no such relationship.`,
  [PARSE_WARNING_CODES.unnumberedParagraph]: (warning) =>
    `Unnumbered ${plural(warning.count, "paragraph")} whose numbering definitions are missing.`,
  [PARSE_WARNING_CODES.unnumberedStyle]: (warning) =>
    `Unnumbered style${quoted(warning.value)} whose numbering definition is missing.`,
  [PARSE_WARNING_CODES.unbalancedMoveRange]: (warning) =>
    `Removed ${plural(warning.count, "unbalanced tracked move range marker")}.`,
  [PARSE_WARNING_CODES.headerFooterTypeOutsideEnum]: (warning) =>
    `Read header/footer type${quoted(warning.value)} as "default"${where(warning)}; ST_HdrFtr is even, default or first.`,
  [PARSE_WARNING_CODES.unrecognisedOnOffValue]: (warning) =>
    `Ignored on/off value${quoted(warning.value)}${where(warning)}; ST_OnOff is 1, 0, true, false, on or off.`,
  [PARSE_WARNING_CODES.borderWithoutValue]: (warning) =>
    `Read a border with no w:val${where(warning)} as having no border style.`,
  [PARSE_WARNING_CODES.styleSetDuplicateStyleId]: (warning) =>
    `Dropped a style repeating the id${quoted(warning.value)} another style in the set already defines.`,
  [PARSE_WARNING_CODES.styleSetInitialStyleMissing]: (warning) =>
    `The style set names initial paragraph style${quoted(warning.value)}, which it does not contain; used the set's default instead${where(warning)}.`,
} as const satisfies Record<ParseWarningCode, ParseWarningFormatter>;

export const formatParseWarning = (warning: ParseWarning): string =>
  PARSE_WARNING_MESSAGES[warning.code](warning);

export const formatParseWarnings = (warnings: readonly ParseWarning[]): string[] =>
  warnings.map(formatParseWarning);
