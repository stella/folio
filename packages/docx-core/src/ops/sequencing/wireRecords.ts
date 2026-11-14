/**
 * Bounded validation of the model records operations embed on the wire.
 *
 * The batch decoder first refuses anything but plain JSON data within its
 * depth and byte bounds. These validators then check every embedded record's
 * shape: each content node's kind is one the model defines for its slot,
 * each record carries only the fields its model type declares, paragraph ids
 * are well formed, scalar formatting values have their model types, and
 * nested content is checked recursively. A field whose value is itself a
 * deep property record (borders, shading, drawings, style and section
 * properties) is bounded by the JSON checks alone. Each table below is
 * exhaustive over its model type, so a field added to the model without a
 * decision here stops this module compiling.
 */
import type {
  AbstractNumbering,
  BlockContent,
  Comment,
  DocumentSettings,
  Endnote,
  Footnote,
  HeaderFooter,
  NumberingDefinitions,
  NumberingInstance,
  Paragraph,
  ParagraphContent,
  ParagraphFormatting,
  ParagraphMarkChange,
  ParagraphPropertyChange,
  RecordField,
  Relationship,
  RunContent,
  SectionProperties,
  StyleDefinitions,
  Table,
  TableCell,
  TableCellFormatting,
  TableFormatting,
  TableRow,
  TableRowFormatting,
  TextFormatting,
  TrackedChangeInfo,
} from "../../model/document";
import { PARAGRAPH_ALIGNMENTS } from "../../model/ooxmlEnumerations.gen";
import type { PackageResourceMedia, SectionViewEntry } from "../types";

export type Validator = (value: unknown) => boolean;
export type Fields = Readonly<Record<string, Validator>>;

export const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
export const string: Validator = (value) => typeof value === "string";
export const nonemptyString: Validator = (value) =>
  typeof value === "string" && value.length > 0;
export const boolean: Validator = (value) => typeof value === "boolean";
export const number: Validator = (value) =>
  typeof value === "number" && Number.isFinite(value);
export const natural: Validator = (value) =>
  typeof value === "number" && Number.isSafeInteger(value) && value >= 0;
export const integer: Validator = (value) =>
  typeof value === "number" && Number.isSafeInteger(value);
export const blockId: Validator = (value) =>
  typeof value === "string" && /^[0-9A-F]{8}$/u.test(value);
export const absent: Validator = (value) => value === undefined;
/** Any value: the decoder already bounds it as plain JSON data. */
export const boundedJson: Validator = () => true;
export const present: Validator = (value) => value !== undefined;
export const optional =
  (check: Validator): Validator =>
  (value) =>
    value === undefined || check(value);
export const nullable =
  (check: Validator): Validator =>
  (value) =>
    value === null || check(value);
export const literal =
  (...values: readonly unknown[]): Validator =>
  (value) =>
    values.includes(value);
export const array =
  (check: Validator): Validator =>
  (value) =>
    Array.isArray(value) && value.every(check);
export const object =
  (fields: Fields): Validator =>
  (value) =>
    isRecord(value) &&
    Object.keys(value).every((key) => Object.hasOwn(fields, key)) &&
    Object.entries(fields).every(([key, check]) => check(value[key]));
/** Every field optional; a patch also admits null, which removes the key. */
export const properties = (fields: Fields, patch = false): Validator =>
  object(
    Object.fromEntries(
      Object.entries(fields).map(([key, check]) => [
        key,
        optional(patch ? nullable(check) : check),
      ]),
    ),
  );

/** Every field of every member of a model union, without the writable mark. */
type AllFields<Value> = Value extends unknown ? RecordField<Value> : never;
/** One validator per declared field: none may be left out. */
type RecordFields<Value> = { readonly [Key in AllFields<Value>]-?: Validator };
/** The members of a union whose discriminant admits `Kind`. */
type KindMember<
  Union extends { type: string },
  Kind extends string,
> = Union extends unknown
  ? Kind extends Union["type"]
    ? Union
    : never
  : never;
type Kinds<Union extends { type: string }> = {
  readonly [Kind in Union["type"]]: RecordFields<KindMember<Union, Kind>>;
};

const kinded =
  (table: Readonly<Record<string, Fields>>): Validator =>
  (value) => {
    if (!isRecord(value) || typeof value["type"] !== "string") return false;
    const fields = Object.entries(table).find(
      ([kind]) => kind === value["type"],
    )?.[1];
    return fields !== undefined && object(fields)(value);
  };

// Recursive slots refer to these through function declarations.
function inline(value: unknown): boolean {
  return inlineContent(value);
}
function block(value: unknown): boolean {
  return blockContent(value);
}

const inlines = array(inline);
const blocks = array(block);

/** Scalar run properties with their model types; the rest stay bounded JSON. */
export const runPropertyFields = {
  bold: boolean,
  boldCs: boolean,
  italic: boolean,
  italicCs: boolean,
  underline: boundedJson,
  strike: boolean,
  doubleStrike: boolean,
  vertAlign: literal("baseline", "superscript", "subscript"),
  smallCaps: boolean,
  allCaps: boolean,
  hidden: boolean,
  noProof: boolean,
  color: boundedJson,
  highlight: boundedJson,
  shading: boundedJson,
  fontSize: number,
  fontSizeCs: number,
  fontFamily: boundedJson,
  language: object({
    val: optional(string),
    eastAsia: optional(string),
    bidi: optional(string),
  }),
  spacing: number,
  position: number,
  scale: number,
  kerning: number,
  effect: boundedJson,
  emphasisMark: boundedJson,
  emboss: boolean,
  imprint: boolean,
  outline: boolean,
  shadow: boolean,
  rtl: boolean,
  cs: boolean,
  styleId: string,
  preserved: boundedJson,
} satisfies RecordFields<TextFormatting>;

/** Scalar paragraph properties with their model types; the rest stay bounded JSON. */
export const paragraphPropertyFields = {
  alignment: literal(...PARAGRAPH_ALIGNMENTS),
  bidi: boolean,
  kinsoku: boolean,
  autoSpaceDE: boolean,
  autoSpaceDN: boolean,
  overflowPunctuation: boolean,
  spaceBefore: number,
  spaceAfter: number,
  lineSpacing: number,
  lineSpacingRule: literal("auto", "exact", "atLeast"),
  snapToGrid: boolean,
  beforeAutospacing: boolean,
  afterAutospacing: boolean,
  spacingExplicit: boundedJson,
  spacingPreservedAttributes: boundedJson,
  indentLeft: number,
  indentRight: number,
  indentFirstLine: number,
  hangingIndent: boolean,
  indentPreservedAttributes: boundedJson,
  borders: boundedJson,
  shading: boundedJson,
  tabs: boundedJson,
  keepNext: boolean,
  keepLines: boolean,
  widowControl: boolean,
  pageBreakBefore: boolean,
  contextualSpacing: boolean,
  numPr: boundedJson,
  numPrFromStyle: boundedJson,
  numberingChangeXml: string,
  numberingInsertionXml: string,
  outlineLevel: boundedJson,
  styleId: string,
  frame: boundedJson,
  suppressLineNumbers: boolean,
  suppressAutoHyphens: boolean,
  runProperties: (value) => runProperties(value),
  runInWithNext: boolean,
  preserved: boundedJson,
} satisfies RecordFields<ParagraphFormatting>;

export const runProperties: Validator = properties(runPropertyFields);
export const paragraphProperties: Validator = properties(
  paragraphPropertyFields,
);

export const trackedChangeInfo: Validator = object({
  id: natural,
  author: string,
  date: optional(string),
  initials: optional(string),
  utcDate: optional(object({ attribute: string, value: string })),
  suggestionId: optional(string),
} satisfies RecordFields<TrackedChangeInfo>);

export const paragraphMarkChange: Validator = object({
  kind: literal("moveFrom", "moveTo", "ins", "del"),
  info: trackedChangeInfo,
  resolutionJoin: optional(natural),
  survivor: optional(literal("first")),
} satisfies RecordFields<ParagraphMarkChange>);

export const paragraphPropertyChange: Validator = object({
  type: literal("paragraphPropertyChange"),
  info: present,
  previousFormatting: optional(paragraphProperties),
  currentFormatting: optional(paragraphProperties),
} satisfies RecordFields<ParagraphPropertyChange>);

const runContentKinds = {
  text: { type: string, text: string },
  tab: { type: string, positional: boundedJson },
  break: {
    type: string,
    sourceElement: optional(literal("br", "cr")),
    breakType: optional(literal("page", "column", "textWrapping")),
    clear: optional(literal("none", "left", "right", "all")),
  },
  symbol: { type: string, font: string, char: string },
  footnoteRef: { type: string, id: natural, customMarkFollows: boundedJson },
  endnoteRef: { type: string, id: natural, customMarkFollows: boundedJson },
  noteMarker: { type: string, kind: literal("footnote", "endnote") },
  fieldChar: {
    type: string,
    charType: literal("begin", "separate", "end"),
    fldLock: optional(boolean),
    dirty: optional(boolean),
    originalValue: optional(string),
  },
  instrText: { type: string, text: string },
  softHyphen: { type: string },
  noBreakHyphen: { type: string },
  renderedPageBreak: { type: string },
  preservedXml: { type: string, xml: string, text: string },
  drawing: {
    type: string,
    image: present,
    rawXml: optional(string),
    rawImageFingerprint: optional(string),
    rawXmlMode: boundedJson,
  },
  shape: { type: string, shape: present, alternateContent: boundedJson },
} satisfies Kinds<RunContent>;

const runContent = kinded(runContentKinds);

const marker = {
  id: natural,
  displacedByCustomXml: boundedJson,
  copyOriginAncestors: boundedJson,
} as const;
const moveRangeStart = {
  type: string,
  ...marker,
  name: string,
  colFirst: optional(natural),
  colLast: optional(natural),
  author: string,
  date: optional(string),
};
const trackedRun = {
  type: string,
  resolutionJoins: boundedJson,
  info: trackedChangeInfo,
  content: inlines,
};

const run = {
  type: string,
  formatting: optional(runProperties),
  propertyChanges: boundedJson,
  content: array(runContent),
  preservedAttributes: boundedJson,
};

const inlineContentKinds = {
  run,
  hyperlink: {
    type: string,
    rId: optional(string),
    href: optional(string),
    anchor: optional(string),
    tooltip: optional(string),
    target: optional(string),
    history: optional(boolean),
    docLocation: optional(string),
    children: inlines,
  },
  bookmarkStart: {
    type: string,
    ...marker,
    name: string,
    colFirst: optional(natural),
    colLast: optional(natural),
  },
  bookmarkEnd: { type: string, ...marker },
  simpleField: {
    type: string,
    instruction: string,
    fieldType: boundedJson,
    content: inlines,
    fldLock: optional(boolean),
    dirty: optional(boolean),
  },
  complexField: {
    type: string,
    instruction: string,
    fieldType: boundedJson,
    fieldCode: inlines,
    fieldResult: inlines,
    formatting: optional(runProperties),
    fldLock: optional(boolean),
    dirty: optional(boolean),
    formFieldDataXml: optional(string),
    fieldResultIsFallback: optional(boolean),
  },
  inlineSdt: { type: string, properties: present, content: inlines },
  commentRangeStart: {
    type: string,
    id: natural,
    displacedByCustomXml: boundedJson,
  },
  commentRangeEnd: {
    type: string,
    id: natural,
    displacedByCustomXml: boundedJson,
  },
  commentReference: { type: string, id: natural },
  insertion: trackedRun,
  deletion: trackedRun,
  moveFrom: trackedRun,
  moveTo: trackedRun,
  moveFromRangeStart: moveRangeStart,
  moveFromRangeEnd: { type: string, ...marker },
  moveToRangeStart: moveRangeStart,
  moveToRangeEnd: { type: string, ...marker },
  inlineWrapper: {
    type: string,
    kind: literal("bidi", "smartTag", "customXml"),
    control: boundedJson,
    direction: optional(literal("ltr", "rtl")),
    element: optional(string),
    uri: optional(string),
    propertiesXml: optional(string),
    content: inlines,
  },
  mathEquation: {
    type: string,
    display: literal("inline", "block"),
    ommlXml: string,
    plainText: optional(string),
  },
  preservedInline: {
    type: string,
    xml: string,
    text: string,
    foldedListNumber: boundedJson,
  },
} satisfies Kinds<ParagraphContent>;

/** One paragraph content record, its nested content checked as well. */
export const inlineContent: Validator = kinded(inlineContentKinds);

const paragraphFields = {
  type: literal("paragraph"),
  paraId: optional(blockId),
  textId: optional(blockId),
  formatting: optional(paragraphProperties),
  propertyChanges: optional(array(paragraphPropertyChange)),
  pPrMark: optional(paragraphMarkChange),
  reviewCarrier: boundedJson,
  content: inlines,
  listRendering: boundedJson,
  renderedPageBreakBefore: optional(boolean),
  sectionProperties: boundedJson,
  preservedAttributes: boundedJson,
} satisfies RecordFields<Paragraph>;

/** A paragraph record carrying a well-formed identity. */
export const paragraph: Validator = object({
  ...paragraphFields,
  paraId: blockId,
});

/** A paragraph record whose identity, when it has one, is well formed. */
export const paragraphRecord: Validator = object(paragraphFields);

/** The fields a split gives its new paragraph, or a join retires. */
export const splitParagraphFields: Validator = object(
  Object.fromEntries(
    Object.entries(paragraphFields).filter(
      ([key]) =>
        key !== "type" &&
        key !== "paraId" &&
        key !== "content" &&
        key !== "sectionProperties" &&
        key !== "pPrMark",
    ),
  ),
);

/** A paragraph's review fields: its properties, their changes and its mark. */
export const paragraphReviewFields: Validator = object({
  formatting: optional(paragraphProperties),
  propertyChanges: optional(array(paragraphPropertyChange)),
  pPrMark: optional(paragraphMarkChange),
});

const tableCellFields = {
  type: literal("tableCell"),
  id: optional(string),
  formatting: boundedJson,
  propertyChanges: boundedJson,
  structuralChange: boundedJson,
  content: blocks,
  contentControls: boundedJson,
  carrierStack: boundedJson,
} satisfies RecordFields<TableCell>;

const tableRowFields = {
  type: literal("tableRow"),
  formatting: boundedJson,
  tablePropertyExceptions: boundedJson,
  propertyChanges: boundedJson,
  tablePropertyExceptionChanges: boundedJson,
  structuralChange: boundedJson,
  cells: array(object(tableCellFields)),
  preserved: boundedJson,
  preservedAttributes: boundedJson,
  bookmarks: boundedJson,
  contentControls: boundedJson,
  carrierStack: boundedJson,
} satisfies RecordFields<TableRow>;

export const tableRow: Validator = object(tableRowFields);

const tableFields = {
  type: string,
  formatting: boundedJson,
  propertyChanges: boundedJson,
  columnWidths: optional(array(number)),
  rows: array(tableRow),
  preserved: boundedJson,
  bookmarks: boundedJson,
  carrierStack: boundedJson,
} satisfies RecordFields<Table>;

const blockContentKinds = {
  paragraph: paragraphFields,
  table: tableFields,
  blockSdt: { type: string, properties: present, content: blocks },
  blockCustomXml: {
    type: string,
    openingXml: string,
    closingXml: string,
    content: blocks,
  },
  preservedBlock: { type: string, xml: string, readerText: optional(string) },
  bookmarkStart: inlineContentKinds.bookmarkStart,
  bookmarkEnd: inlineContentKinds.bookmarkEnd,
  commentRangeStart: inlineContentKinds.commentRangeStart,
  commentRangeEnd: inlineContentKinds.commentRangeEnd,
  moveFromRangeStart: moveRangeStart,
  moveFromRangeEnd: inlineContentKinds.moveFromRangeEnd,
  moveToRangeStart: moveRangeStart,
  moveToRangeEnd: inlineContentKinds.moveToRangeEnd,
} satisfies Kinds<BlockContent>;

/** One block record, its nested blocks and paragraphs checked as well. */
export const blockContent: Validator = kinded(blockContentKinds);
export const table: Validator = kinded({ table: tableFields });

export const sectionPropertyFields = {
  pageWidth: number,
  pageHeight: number,
  orientation: boundedJson,
  marginTop: number,
  marginBottom: number,
  marginLeft: number,
  marginRight: number,
  headerDistance: number,
  footerDistance: number,
  gutter: number,
  columnCount: number,
  columnSpace: number,
  equalWidth: boolean,
  separator: boolean,
  columns: boundedJson,
  sectionStart: boundedJson,
  verticalAlign: boundedJson,
  textDirection: boundedJson,
  bidi: boolean,
  headerReferences: boundedJson,
  footerReferences: boundedJson,
  titlePg: boolean,
  evenAndOddHeaders: boolean,
  lineNumbers: boundedJson,
  pageNumbering: boundedJson,
  pageBorders: boundedJson,
  background: boundedJson,
  footnotePr: boundedJson,
  footnoteColumns: boundedJson,
  endnotePr: boundedJson,
  docGrid: boundedJson,
  paperSrcFirst: boundedJson,
  paperSrcOther: boundedJson,
  formProtection: boundedJson,
  noEndnote: boundedJson,
  rtlGutter: boundedJson,
  printerSettingsRelationshipId: optional(string),
  propertyChanges: boundedJson,
  preserved: boundedJson,
  preservedAttributes: boundedJson,
} satisfies RecordFields<SectionProperties>;

export const sectionProperties: Validator = properties(sectionPropertyFields);

/** Table property keys; their values stay bounded JSON. */
export const tableFormattingFields = {
  width: boundedJson,
  justification: boundedJson,
  cellSpacing: boundedJson,
  indent: boundedJson,
  borders: boundedJson,
  cellMargins: boundedJson,
  layout: boundedJson,
  styleId: boundedJson,
  look: boundedJson,
  shading: boundedJson,
  overlap: boundedJson,
  floating: boundedJson,
  bidi: boundedJson,
  caption: boundedJson,
  description: boundedJson,
  rowBandSize: boundedJson,
  columnBandSize: boundedJson,
  preserved: boundedJson,
  gridSourceXml: boundedJson,
  gridChange: boundedJson,
  sourceXml: boundedJson,
} satisfies RecordFields<TableFormatting>;

export const tableCellFormattingFields = {
  width: boundedJson,
  borders: boundedJson,
  margins: boundedJson,
  shading: boundedJson,
  verticalAlign: boundedJson,
  textDirection: boundedJson,
  gridSpan: boundedJson,
  vMerge: boundedJson,
  fitText: boundedJson,
  noWrap: boundedJson,
  hideMark: boundedJson,
  conditionalFormat: boundedJson,
  preserved: boundedJson,
  sourceXml: boundedJson,
} satisfies RecordFields<TableCellFormatting>;

export const tableRowFormattingFields = {
  gridBefore: boundedJson,
  widthBefore: boundedJson,
  gridAfter: boundedJson,
  widthAfter: boundedJson,
  height: boundedJson,
  heightRule: boundedJson,
  header: boundedJson,
  cantSplit: boundedJson,
  justification: boundedJson,
  hidden: boundedJson,
  conditionalFormat: boundedJson,
  preserved: boundedJson,
  sourceXml: boundedJson,
} satisfies RecordFields<TableRowFormatting>;

const headerFooterType = literal("default", "first", "even");

export const headerFooter: Validator = object({
  type: literal("header", "footer"),
  hdrFtrType: headerFooterType,
  content: blocks,
  watermark: boundedJson,
  rawWatermarkXml: optional(string),
  verbatimXml: optional(string),
  verbatimFingerprint: optional(string),
  watermarkBlockIndex: optional(natural),
} satisfies RecordFields<HeaderFooter>);

const noteType = optional(
  literal("normal", "separator", "continuationSeparator", "continuationNotice"),
);
export const note: Validator = kinded({
  footnote: {
    type: string,
    id: integer,
    noteType,
    content: blocks,
  } satisfies RecordFields<Footnote>,
  endnote: {
    type: string,
    id: integer,
    noteType,
    content: blocks,
  } satisfies RecordFields<Endnote>,
});

/** Map entries as an operation carries them: `[key, value]` pairs. */
export const mapEntries = (key: Validator, value: Validator): Validator =>
  array(
    (entry) =>
      Array.isArray(entry) &&
      entry.length === 2 &&
      key(entry[0]) &&
      value(entry[1]),
  );

const sectionMapState = (value: unknown): boolean =>
  isRecord(value) &&
  (object({ type: literal("omitted", "undefined") })(value) ||
    object({
      type: literal("entries"),
      value: mapEntries(headerFooterType, headerFooter),
    })(value));

export const sectionViewEntry: Validator = object({
  properties: sectionProperties,
  headers: sectionMapState,
  footers: sectionMapState,
} satisfies RecordFields<SectionViewEntry>);

export const sectionViewChange: Validator = object({
  expected: array(sectionViewEntry),
  restore: array(sectionViewEntry),
});

/** A tagged presence: omitted, own undefined, or a value. */
export const presence = (tag: string, value: Validator): Validator =>
  literalOr(object({ type: literal(tag), value }));
const literalOr =
  (valued: Validator): Validator =>
  (value) =>
    object({ type: literal("omitted", "undefined") })(value) || valued(value);

export const sectionViewState: Validator = literalOr(
  object({ type: literal("sections"), value: array(sectionViewEntry) }),
);

export const numberingInstance: Validator = object({
  numId: integer,
  abstractNumId: integer,
  levelOverrides: boundedJson,
  preserved: boundedJson,
  preservedAttributes: boundedJson,
} satisfies RecordFields<NumberingInstance>);

export const abstractNumbering: Validator = object({
  abstractNumId: integer,
  nsid: optional(string),
  multiLevelType: boundedJson,
  tmpl: optional(string),
  numStyleLink: optional(string),
  styleLink: optional(string),
  levels: array(isRecord),
  name: optional(string),
  preserved: boundedJson,
  preservedAttributes: boundedJson,
} satisfies RecordFields<AbstractNumbering>);

export const numberingDefinitions: Validator = object({
  abstractNums: array(abstractNumbering),
  nums: array(numberingInstance),
  preserved: boundedJson,
  preservedAttributes: boundedJson,
  danglingNumIds: optional(array(integer)),
} satisfies RecordFields<NumberingDefinitions>);

export const styleDefinitions: Validator = object({
  docDefaults: boundedJson,
  latentStyles: boundedJson,
  styles: array(isRecord),
} satisfies RecordFields<StyleDefinitions>);

export const relationship: Validator = object({
  id: nonemptyString,
  type: string,
  target: string,
  targetMode: optional(literal("External", "Internal")),
} satisfies RecordFields<Relationship>);

export const packageResourceMedia: Validator = object({
  path: nonemptyString,
  filename: optional(string),
  mimeType: string,
  base64: optional(string),
  dataUrl: optional(string),
  data: array(
    (byte) =>
      typeof byte === "number" &&
      Number.isInteger(byte) &&
      byte >= 0 &&
      byte <= 255,
  ),
} satisfies RecordFields<PackageResourceMedia>);

export const documentSettings: Validator = object({
  compatibilityMode: boundedJson,
  noColumnBalance: optional(boolean),
  defaultTabStop: number,
  evenAndOddHeaders: optional(boolean),
  mirrorMargins: optional(boolean),
  adjustLineHeightInTable: optional(boolean),
  doNotUseIndentAsNumberingTabStop: optional(boolean),
  updateFields: optional(boolean),
  themeFontLang: boundedJson,
  autoHyphenation: optional(boolean),
  doNotHyphenateCaps: optional(boolean),
  consecutiveHyphenLimit: optional(number),
  hyphenationZoneTwips: optional(number),
  lineBreakRules: boundedJson,
  splitPageBreakAndParagraphMark: optional(boolean),
  endnotePr: boundedJson,
} satisfies RecordFields<DocumentSettings>);

const commentFields = {
  id: natural,
  anchorKind: optional(literal("point", "range")),
  author: string,
  initials: optional(string),
  date: optional(string),
  // A loaded comment's paragraphs may carry no identity of their own.
  content: array(paragraphRecord),
  annotationReferenceFormatting: optional(runProperties),
  parentId: optional(natural),
  done: optional(boolean),
  suggestionId: optional(string),
  preserved: boundedJson,
} satisfies RecordFields<Comment>;

/** A comment record as it stands in the comment list. */
export const comment: Validator = object(commentFields);

/**
 * A comment as a creation or a content edit carries it: every paragraph has
 * its own identity, and the thread a reply joins belongs to the anchor.
 */
export const newComment: Validator = object({
  ...Object.fromEntries(
    Object.entries(commentFields).filter(([key]) => key !== "parentId"),
  ),
  content: array(paragraph),
});
