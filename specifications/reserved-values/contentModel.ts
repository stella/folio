/** Reserved-value decisions for the content model. */

import type { ExhaustiveFields } from "../../packages/docx-core/src/model/exhaustiveFields";
import type {
  BlockSdt,
  BookmarkEnd,
  BookmarkStart,
  BreakContent,
  Comment,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  ComplexField,
  Deletion,
  DocumentBody,
  DrawingContent,
  FieldCharContent,
  FooterReference,
  HeaderFooter,
  HeaderReference,
  Hyperlink,
  Image,
  ImageCrop,
  ImageFrameLocks,
  ImagePadding,
  ImagePosition,
  ImageSize,
  ImageTransform,
  ImageWrap,
  InlineSdt,
  Insertion,
  InstrTextContent,
  MathEquation,
  MoveFrom,
  MoveFromRangeEnd,
  MoveFromRangeStart,
  MoveTo,
  MoveToRangeEnd,
  MoveToRangeStart,
  NoBreakHyphenContent,
  NoteReferenceContent,
  Paragraph,
  ParagraphMarkChange,
  ParagraphPropertyChange,
  PictureWatermark,
  PositionalTab,
  PropertyChangeInfo,
  RenderedPageBreakContent,
  Run,
  RunPropertyChange,
  SdtProperties,
  Section,
  SectionPropertyChange,
  Shape,
  ShapeContent,
  ShapeFill,
  ShapeGeometryAdjustment,
  ShapeOutline,
  ShapeTextBody,
  SimpleField,
  SoftHyphenContent,
  SymbolContent,
  TabContent,
  Table,
  TableCell,
  TableCellPropertyChange,
  TablePropertyChange,
  TableRow,
  TableRowPropertyChange,
  TableStructuralChangeInfo,
  TextBox,
  TextContent,
  TextWatermark,
  TrackedChangeInfo,
} from "../../packages/docx-core/src/model/content";
import {
  NO_RESERVED_VALUE,
  notModelled,
  readerOwned,
  type ReservedValueDisposition,
  toggle,
  type UnionFields,
} from "./disposition";
import { RESERVED_VALUE_READERS } from "./readers";

export const TEXT_CONTENT_RESERVED = {
  type: NO_RESERVED_VALUE,
  text: NO_RESERVED_VALUE,
} satisfies Record<keyof TextContent, ReservedValueDisposition>;

export type ExhaustiveTextContentReserved = ExhaustiveFields<
  TextContent,
  keyof typeof TEXT_CONTENT_RESERVED
>;

export const POSITIONAL_TAB_RESERVED = {
  relativeTo: NO_RESERVED_VALUE,
  alignment: NO_RESERVED_VALUE,
  leader: readerOwned({
    slot: "w:ptab@leader",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.runContent,
  }),
} satisfies Record<keyof PositionalTab, ReservedValueDisposition>;

export type ExhaustivePositionalTabReserved = ExhaustiveFields<
  PositionalTab,
  keyof typeof POSITIONAL_TAB_RESERVED
>;

export const TAB_CONTENT_RESERVED = {
  type: NO_RESERVED_VALUE,
  positional: NO_RESERVED_VALUE,
} satisfies Record<keyof TabContent, ReservedValueDisposition>;

export type ExhaustiveTabContentReserved = ExhaustiveFields<
  TabContent,
  keyof typeof TAB_CONTENT_RESERVED
>;

export const BREAK_CONTENT_RESERVED = {
  type: NO_RESERVED_VALUE,
  breakType: NO_RESERVED_VALUE,
  clear: readerOwned({
    slot: "w:br@clear",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.runContent,
  }),
} satisfies Record<keyof BreakContent, ReservedValueDisposition>;

export type ExhaustiveBreakContentReserved = ExhaustiveFields<
  BreakContent,
  keyof typeof BREAK_CONTENT_RESERVED
>;

export const SYMBOL_CONTENT_RESERVED = {
  type: NO_RESERVED_VALUE,
  font: NO_RESERVED_VALUE,
  char: NO_RESERVED_VALUE,
} satisfies Record<keyof SymbolContent, ReservedValueDisposition>;

export type ExhaustiveSymbolContentReserved = ExhaustiveFields<
  SymbolContent,
  keyof typeof SYMBOL_CONTENT_RESERVED
>;

export const NOTE_REFERENCE_CONTENT_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: readerOwned({
    slot: "w:footnote@id|w:endnote@id|w:footnoteReference@id|w:endnoteReference@id",
    sentinel: "-1|0",
    reader: RESERVED_VALUE_READERS.noteType,
    evidence: "note-ids-minus-one-and-zero-are-reserved",
  }),
} satisfies Record<keyof NoteReferenceContent, ReservedValueDisposition>;

export type ExhaustiveNoteReferenceContentReserved = ExhaustiveFields<
  NoteReferenceContent,
  keyof typeof NOTE_REFERENCE_CONTENT_RESERVED
>;

export const FIELD_CHAR_CONTENT_RESERVED = {
  type: NO_RESERVED_VALUE,
  charType: NO_RESERVED_VALUE,
  fldLock: toggle("w:fldChar@fldLock"),
  dirty: toggle("w:fldChar@dirty"),
  originalValue: NO_RESERVED_VALUE,
} satisfies Record<keyof FieldCharContent, ReservedValueDisposition>;

export type ExhaustiveFieldCharContentReserved = ExhaustiveFields<
  FieldCharContent,
  keyof typeof FIELD_CHAR_CONTENT_RESERVED
>;

export const INSTR_TEXT_CONTENT_RESERVED = {
  type: NO_RESERVED_VALUE,
  text: NO_RESERVED_VALUE,
} satisfies Record<keyof InstrTextContent, ReservedValueDisposition>;

export type ExhaustiveInstrTextContentReserved = ExhaustiveFields<
  InstrTextContent,
  keyof typeof INSTR_TEXT_CONTENT_RESERVED
>;

export const SOFT_HYPHEN_CONTENT_RESERVED = {
  type: NO_RESERVED_VALUE,
} satisfies Record<keyof SoftHyphenContent, ReservedValueDisposition>;

export type ExhaustiveSoftHyphenContentReserved = ExhaustiveFields<
  SoftHyphenContent,
  keyof typeof SOFT_HYPHEN_CONTENT_RESERVED
>;

export const NO_BREAK_HYPHEN_CONTENT_RESERVED = {
  type: NO_RESERVED_VALUE,
} satisfies Record<keyof NoBreakHyphenContent, ReservedValueDisposition>;

export type ExhaustiveNoBreakHyphenContentReserved = ExhaustiveFields<
  NoBreakHyphenContent,
  keyof typeof NO_BREAK_HYPHEN_CONTENT_RESERVED
>;

export const RENDERED_PAGE_BREAK_CONTENT_RESERVED = {
  type: NO_RESERVED_VALUE,
} satisfies Record<keyof RenderedPageBreakContent, ReservedValueDisposition>;

export type ExhaustiveRenderedPageBreakContentReserved = ExhaustiveFields<
  RenderedPageBreakContent,
  keyof typeof RENDERED_PAGE_BREAK_CONTENT_RESERVED
>;

export const DRAWING_CONTENT_RESERVED = {
  type: NO_RESERVED_VALUE,
  image: NO_RESERVED_VALUE,
  rawXml: NO_RESERVED_VALUE,
  rawImageFingerprint: NO_RESERVED_VALUE,
  rawXmlMode: NO_RESERVED_VALUE,
} satisfies Record<UnionFields<DrawingContent>, ReservedValueDisposition>;

export const SHAPE_CONTENT_RESERVED = {
  type: NO_RESERVED_VALUE,
  shape: NO_RESERVED_VALUE,
} satisfies Record<keyof ShapeContent, ReservedValueDisposition>;

export type ExhaustiveShapeContentReserved = ExhaustiveFields<
  ShapeContent,
  keyof typeof SHAPE_CONTENT_RESERVED
>;

export const RUN_RESERVED = {
  type: NO_RESERVED_VALUE,
  formatting: NO_RESERVED_VALUE,
  propertyChanges: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
  // Attributes replayed as the source wrote them. A reserved value is a
  // spelling the model interprets; this slot interprets nothing.
  preservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof Run, ReservedValueDisposition>;

export type ExhaustiveRunReserved = ExhaustiveFields<Run, keyof typeof RUN_RESERVED>;

export const HYPERLINK_RESERVED = {
  type: NO_RESERVED_VALUE,
  rId: NO_RESERVED_VALUE,
  href: NO_RESERVED_VALUE,
  anchor: NO_RESERVED_VALUE,
  tooltip: NO_RESERVED_VALUE,
  target: NO_RESERVED_VALUE,
  history: toggle("w:hyperlink@history"),
  docLocation: NO_RESERVED_VALUE,
  children: NO_RESERVED_VALUE,
} satisfies Record<keyof Hyperlink, ReservedValueDisposition>;

export type ExhaustiveHyperlinkReserved = ExhaustiveFields<
  Hyperlink,
  keyof typeof HYPERLINK_RESERVED
>;

export const BOOKMARK_START_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  name: NO_RESERVED_VALUE,
  colFirst: NO_RESERVED_VALUE,
  colLast: NO_RESERVED_VALUE,
  displacedByCustomXml: NO_RESERVED_VALUE,
} satisfies Record<keyof BookmarkStart, ReservedValueDisposition>;

export type ExhaustiveBookmarkStartReserved = ExhaustiveFields<
  BookmarkStart,
  keyof typeof BOOKMARK_START_RESERVED
>;

export const BOOKMARK_END_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  displacedByCustomXml: NO_RESERVED_VALUE,
} satisfies Record<keyof BookmarkEnd, ReservedValueDisposition>;

export type ExhaustiveBookmarkEndReserved = ExhaustiveFields<
  BookmarkEnd,
  keyof typeof BOOKMARK_END_RESERVED
>;

export const SIMPLE_FIELD_RESERVED = {
  type: NO_RESERVED_VALUE,
  instruction: NO_RESERVED_VALUE,
  fieldType: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
  fldLock: toggle("w:fldSimple@fldLock"),
  dirty: toggle("w:fldSimple@dirty"),
} satisfies Record<keyof SimpleField, ReservedValueDisposition>;

export type ExhaustiveSimpleFieldReserved = ExhaustiveFields<
  SimpleField,
  keyof typeof SIMPLE_FIELD_RESERVED
>;

export const COMPLEX_FIELD_RESERVED = {
  type: NO_RESERVED_VALUE,
  instruction: NO_RESERVED_VALUE,
  fieldType: NO_RESERVED_VALUE,
  fieldCode: NO_RESERVED_VALUE,
  fieldResult: NO_RESERVED_VALUE,
  formatting: NO_RESERVED_VALUE,
  fldLock: toggle("w:fldChar@fldLock"),
  dirty: toggle("w:fldChar@dirty"),
} satisfies Record<keyof ComplexField, ReservedValueDisposition>;

export type ExhaustiveComplexFieldReserved = ExhaustiveFields<
  ComplexField,
  keyof typeof COMPLEX_FIELD_RESERVED
>;

export const IMAGE_SIZE_RESERVED = {
  width: NO_RESERVED_VALUE,
  height: NO_RESERVED_VALUE,
} satisfies Record<keyof ImageSize, ReservedValueDisposition>;

export type ExhaustiveImageSizeReserved = ExhaustiveFields<
  ImageSize,
  keyof typeof IMAGE_SIZE_RESERVED
>;

export const IMAGE_WRAP_RESERVED = {
  type: readerOwned({
    slot: "wp:anchor@behindDoc",
    sentinel: "0|false|off",
    reader: RESERVED_VALUE_READERS.behindDoc,
  }),
  wrapText: NO_RESERVED_VALUE,
  distT: NO_RESERVED_VALUE,
  distB: NO_RESERVED_VALUE,
  distL: NO_RESERVED_VALUE,
  distR: NO_RESERVED_VALUE,
} satisfies Record<keyof ImageWrap, ReservedValueDisposition>;

export type ExhaustiveImageWrapReserved = ExhaustiveFields<
  ImageWrap,
  keyof typeof IMAGE_WRAP_RESERVED
>;

export const IMAGE_POSITION_RESERVED = {
  horizontal: NO_RESERVED_VALUE,
  vertical: NO_RESERVED_VALUE,
} satisfies Record<keyof ImagePosition, ReservedValueDisposition>;

export type ExhaustiveImagePositionReserved = ExhaustiveFields<
  ImagePosition,
  keyof typeof IMAGE_POSITION_RESERVED
>;

export const IMAGE_POSITION_AXIS_RESERVED = {
  relativeTo: NO_RESERVED_VALUE,
  alignment: NO_RESERVED_VALUE,
  posOffset: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<ImagePosition["horizontal"]>, ReservedValueDisposition>;

export type ExhaustiveImagePositionAxisReserved = ExhaustiveFields<
  NonNullable<ImagePosition["horizontal"]>,
  keyof typeof IMAGE_POSITION_AXIS_RESERVED
>;

export const IMAGE_TRANSFORM_RESERVED = {
  rotation: NO_RESERVED_VALUE,
  flipH: NO_RESERVED_VALUE,
  flipV: NO_RESERVED_VALUE,
} satisfies Record<keyof ImageTransform, ReservedValueDisposition>;

export type ExhaustiveImageTransformReserved = ExhaustiveFields<
  ImageTransform,
  keyof typeof IMAGE_TRANSFORM_RESERVED
>;

export const IMAGE_PADDING_RESERVED = {
  top: NO_RESERVED_VALUE,
  bottom: NO_RESERVED_VALUE,
  left: NO_RESERVED_VALUE,
  right: NO_RESERVED_VALUE,
} satisfies Record<keyof ImagePadding, ReservedValueDisposition>;

export type ExhaustiveImagePaddingReserved = ExhaustiveFields<
  ImagePadding,
  keyof typeof IMAGE_PADDING_RESERVED
>;

export const IMAGE_CROP_RESERVED = {
  left: NO_RESERVED_VALUE,
  top: NO_RESERVED_VALUE,
  right: NO_RESERVED_VALUE,
  bottom: NO_RESERVED_VALUE,
} satisfies Record<keyof ImageCrop, ReservedValueDisposition>;

export type ExhaustiveImageCropReserved = ExhaustiveFields<
  ImageCrop,
  keyof typeof IMAGE_CROP_RESERVED
>;

export const IMAGE_FRAME_LOCKS_RESERVED = {
  noGrp: toggle("a:graphicFrameLocks@noGrp"),
  noDrilldown: toggle("a:graphicFrameLocks@noDrilldown"),
  noSelect: toggle("a:graphicFrameLocks@noSelect"),
  noChangeAspect: toggle("a:graphicFrameLocks@noChangeAspect"),
  noMove: toggle("a:graphicFrameLocks@noMove"),
  noResize: toggle("a:graphicFrameLocks@noResize"),
} satisfies Record<keyof ImageFrameLocks, ReservedValueDisposition>;

export type ExhaustiveImageFrameLocksReserved = ExhaustiveFields<
  ImageFrameLocks,
  keyof typeof IMAGE_FRAME_LOCKS_RESERVED
>;

export const IMAGE_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  rId: NO_RESERVED_VALUE,
  src: NO_RESERVED_VALUE,
  mimeType: NO_RESERVED_VALUE,
  filename: NO_RESERVED_VALUE,
  docPrName: NO_RESERVED_VALUE,
  alt: NO_RESERVED_VALUE,
  title: NO_RESERVED_VALUE,
  size: NO_RESERVED_VALUE,
  originalSize: NO_RESERVED_VALUE,
  wrap: NO_RESERVED_VALUE,
  position: NO_RESERVED_VALUE,
  transform: NO_RESERVED_VALUE,
  padding: NO_RESERVED_VALUE,
  crop: NO_RESERVED_VALUE,
  frameLocks: NO_RESERVED_VALUE,
  opacity: readerOwned({
    slot: "a:alphaModFix@amt",
    sentinel: "100%",
    reader: RESERVED_VALUE_READERS.drawing,
  }),
  layoutInCell: toggle("wp:anchor@layoutInCell"),
  allowOverlap: toggle("wp:anchor@allowOverlap"),
  locked: toggle("wp:anchor@locked"),
  anchorHidden: toggle("wp:anchor@hidden"),
  useSimplePosition: toggle("wp:anchor@simplePos"),
  relativeHeight: NO_RESERVED_VALUE,
  simplePosition: NO_RESERVED_VALUE,
  decorative: NO_RESERVED_VALUE,
  hidden: NO_RESERVED_VALUE,
  docPrExtensions: NO_RESERVED_VALUE,
  hlinkHref: NO_RESERVED_VALUE,
  hlinkRId: NO_RESERVED_VALUE,
  hlinkClickSource: NO_RESERVED_VALUE,
  hlinkHoverXml: NO_RESERVED_VALUE,
  outline: NO_RESERVED_VALUE,
  effects: NO_RESERVED_VALUE,
} satisfies Record<keyof Image, ReservedValueDisposition>;

export type ExhaustiveImageReserved = ExhaustiveFields<Image, keyof typeof IMAGE_RESERVED>;

export const IMAGE_EFFECTS_RESERVED = {
  brightness: NO_RESERVED_VALUE,
  contrast: NO_RESERVED_VALUE,
  saturation: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<Image["effects"]>, ReservedValueDisposition>;

export type ExhaustiveImageEffectsReserved = ExhaustiveFields<
  NonNullable<Image["effects"]>,
  keyof typeof IMAGE_EFFECTS_RESERVED
>;

export const SHAPE_GEOMETRY_ADJUSTMENT_RESERVED = {
  name: NO_RESERVED_VALUE,
  formula: NO_RESERVED_VALUE,
} satisfies Record<keyof ShapeGeometryAdjustment, ReservedValueDisposition>;

export type ExhaustiveShapeGeometryAdjustmentReserved = ExhaustiveFields<
  ShapeGeometryAdjustment,
  keyof typeof SHAPE_GEOMETRY_ADJUSTMENT_RESERVED
>;

export const SHAPE_FILL_RESERVED = {
  type: readerOwned({
    slot: "a:noFill",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.shape,
  }),
  rawXml: NO_RESERVED_VALUE,
  color: NO_RESERVED_VALUE,
  gradient: NO_RESERVED_VALUE,
} satisfies Record<keyof ShapeFill, ReservedValueDisposition>;

export type ExhaustiveShapeFillReserved = ExhaustiveFields<
  ShapeFill,
  keyof typeof SHAPE_FILL_RESERVED
>;

export const SHAPE_FILL_GRADIENT_RESERVED = {
  type: NO_RESERVED_VALUE,
  angle: NO_RESERVED_VALUE,
  stops: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<ShapeFill["gradient"]>, ReservedValueDisposition>;

export type ExhaustiveShapeFillGradientReserved = ExhaustiveFields<
  NonNullable<ShapeFill["gradient"]>,
  keyof typeof SHAPE_FILL_GRADIENT_RESERVED
>;

export const SHAPE_FILL_GRADIENT_STOP_RESERVED = {
  position: NO_RESERVED_VALUE,
  color: NO_RESERVED_VALUE,
} satisfies Record<
  keyof NonNullable<ShapeFill["gradient"]>["stops"][number],
  ReservedValueDisposition
>;

export type ExhaustiveShapeFillGradientStopReserved = ExhaustiveFields<
  NonNullable<ShapeFill["gradient"]>["stops"][number],
  keyof typeof SHAPE_FILL_GRADIENT_STOP_RESERVED
>;

export const SHAPE_OUTLINE_RESERVED = {
  rawXml: NO_RESERVED_VALUE,
  width: notModelled({
    slot: "a:ln@w",
    sentinel: "0",
    reason:
      "A zero `@w` is a hairline in OOXML, the thinnest line the renderer can draw. folio maps the width straight to an SVG `stroke-width`, where 0 paints nothing, so a hairline outline disappears instead of thinning.",
  }),
  color: NO_RESERVED_VALUE,
  style: NO_RESERVED_VALUE,
  cap: NO_RESERVED_VALUE,
  join: NO_RESERVED_VALUE,
  headEnd: NO_RESERVED_VALUE,
  tailEnd: NO_RESERVED_VALUE,
} satisfies Record<keyof ShapeOutline, ReservedValueDisposition>;

export type ExhaustiveShapeOutlineReserved = ExhaustiveFields<
  ShapeOutline,
  keyof typeof SHAPE_OUTLINE_RESERVED
>;

export const SHAPE_OUTLINE_END_RESERVED = {
  type: readerOwned({
    slot: "a:headEnd@type|a:tailEnd@type",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.shape,
  }),
  width: NO_RESERVED_VALUE,
  length: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<ShapeOutline["headEnd"]>, ReservedValueDisposition>;

export type ExhaustiveShapeOutlineEndReserved = ExhaustiveFields<
  NonNullable<ShapeOutline["headEnd"]>,
  keyof typeof SHAPE_OUTLINE_END_RESERVED
>;

export const SHAPE_TEXT_BODY_RESERVED = {
  wordArt: NO_RESERVED_VALUE,
  vertical: NO_RESERVED_VALUE,
  rotation: NO_RESERVED_VALUE,
  anchor: NO_RESERVED_VALUE,
  anchorCenter: NO_RESERVED_VALUE,
  autoFit: readerOwned({
    slot: "a:noAutofit",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.shape,
  }),
  textWrap: readerOwned({
    slot: "a:bodyPr@wrap",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.shape,
  }),
  margins: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
} satisfies Record<keyof ShapeTextBody, ReservedValueDisposition>;

export type ExhaustiveShapeTextBodyReserved = ExhaustiveFields<
  ShapeTextBody,
  keyof typeof SHAPE_TEXT_BODY_RESERVED
>;

export const SHAPE_TEXT_BODY_WORD_ART_RESERVED = {
  fromWordArt: NO_RESERVED_VALUE,
  preset: NO_RESERVED_VALUE,
  adjustments: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<ShapeTextBody["wordArt"]>, ReservedValueDisposition>;

export type ExhaustiveShapeTextBodyWordArtReserved = ExhaustiveFields<
  NonNullable<ShapeTextBody["wordArt"]>,
  keyof typeof SHAPE_TEXT_BODY_WORD_ART_RESERVED
>;

export const SHAPE_TEXT_BODY_MARGINS_RESERVED = {
  top: NO_RESERVED_VALUE,
  bottom: NO_RESERVED_VALUE,
  left: NO_RESERVED_VALUE,
  right: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<ShapeTextBody["margins"]>, ReservedValueDisposition>;

export type ExhaustiveShapeTextBodyMarginsReserved = ExhaustiveFields<
  NonNullable<ShapeTextBody["margins"]>,
  keyof typeof SHAPE_TEXT_BODY_MARGINS_RESERVED
>;

export const SHAPE_RESERVED = {
  type: NO_RESERVED_VALUE,
  shapeType: NO_RESERVED_VALUE,
  geometryAdjustments: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  name: NO_RESERVED_VALUE,
  alt: NO_RESERVED_VALUE,
  title: NO_RESERVED_VALUE,
  size: NO_RESERVED_VALUE,
  position: NO_RESERVED_VALUE,
  wrap: NO_RESERVED_VALUE,
  fill: NO_RESERVED_VALUE,
  outline: NO_RESERVED_VALUE,
  transform: NO_RESERVED_VALUE,
  textBody: NO_RESERVED_VALUE,
  customGeometry: NO_RESERVED_VALUE,
} satisfies Record<keyof Shape, ReservedValueDisposition>;

export type ExhaustiveShapeReserved = ExhaustiveFields<Shape, keyof typeof SHAPE_RESERVED>;

export const TEXT_BOX_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  name: NO_RESERVED_VALUE,
  alt: NO_RESERVED_VALUE,
  title: NO_RESERVED_VALUE,
  size: NO_RESERVED_VALUE,
  position: NO_RESERVED_VALUE,
  wrap: NO_RESERVED_VALUE,
  fill: NO_RESERVED_VALUE,
  outline: NO_RESERVED_VALUE,
  transform: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
  wordArt: NO_RESERVED_VALUE,
  autoFit: readerOwned({
    slot: "a:noAutofit",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.textBox,
  }),
  textWrap: readerOwned({
    slot: "a:bodyPr@wrap",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.textBox,
  }),
  verticalAlign: NO_RESERVED_VALUE,
  margins: NO_RESERVED_VALUE,
} satisfies Record<keyof TextBox, ReservedValueDisposition>;

export type ExhaustiveTextBoxReserved = ExhaustiveFields<TextBox, keyof typeof TEXT_BOX_RESERVED>;

export const TEXT_BOX_MARGINS_RESERVED = {
  top: NO_RESERVED_VALUE,
  bottom: NO_RESERVED_VALUE,
  left: NO_RESERVED_VALUE,
  right: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<TextBox["margins"]>, ReservedValueDisposition>;

export type ExhaustiveTextBoxMarginsReserved = ExhaustiveFields<
  NonNullable<TextBox["margins"]>,
  keyof typeof TEXT_BOX_MARGINS_RESERVED
>;

export const TABLE_CELL_RESERVED = {
  type: NO_RESERVED_VALUE,
  formatting: NO_RESERVED_VALUE,
  propertyChanges: NO_RESERVED_VALUE,
  structuralChange: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
} satisfies Record<keyof TableCell, ReservedValueDisposition>;

export type ExhaustiveTableCellReserved = ExhaustiveFields<
  TableCell,
  keyof typeof TABLE_CELL_RESERVED
>;

export const TABLE_ROW_RESERVED = {
  type: NO_RESERVED_VALUE,
  formatting: NO_RESERVED_VALUE,
  propertyChanges: NO_RESERVED_VALUE,
  structuralChange: NO_RESERVED_VALUE,
  cells: NO_RESERVED_VALUE,
  // Captured bytes, replayed as written. A reserved value is a spelling the
  // model interprets; this slot interprets nothing.
  preserved: NO_RESERVED_VALUE,
  preservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof TableRow, ReservedValueDisposition>;

export type ExhaustiveTableRowReserved = ExhaustiveFields<
  TableRow,
  keyof typeof TABLE_ROW_RESERVED
>;

export const TABLE_RESERVED = {
  type: NO_RESERVED_VALUE,
  formatting: NO_RESERVED_VALUE,
  propertyChanges: NO_RESERVED_VALUE,
  columnWidths: NO_RESERVED_VALUE,
  rows: NO_RESERVED_VALUE,
  // Captured bytes, replayed as written. A reserved value is a spelling the
  // model interprets; this slot interprets nothing.
  preserved: NO_RESERVED_VALUE,
} satisfies Record<keyof Table, ReservedValueDisposition>;

export type ExhaustiveTableReserved = ExhaustiveFields<Table, keyof typeof TABLE_RESERVED>;

export const COMMENT_RESERVED = {
  id: NO_RESERVED_VALUE,
  author: NO_RESERVED_VALUE,
  initials: NO_RESERVED_VALUE,
  date: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
  annotationReferenceFormatting: NO_RESERVED_VALUE,
  parentId: NO_RESERVED_VALUE,
  done: NO_RESERVED_VALUE,
  // Captured bytes, replayed as written. A reserved value is a spelling the
  // model interprets; this slot interprets nothing.
  preserved: NO_RESERVED_VALUE,
} satisfies Record<keyof Comment, ReservedValueDisposition>;

export type ExhaustiveCommentReserved = ExhaustiveFields<Comment, keyof typeof COMMENT_RESERVED>;

export const COMMENT_RANGE_START_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  displacedByCustomXml: NO_RESERVED_VALUE,
} satisfies Record<keyof CommentRangeStart, ReservedValueDisposition>;

export type ExhaustiveCommentRangeStartReserved = ExhaustiveFields<
  CommentRangeStart,
  keyof typeof COMMENT_RANGE_START_RESERVED
>;

export const COMMENT_RANGE_END_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  displacedByCustomXml: NO_RESERVED_VALUE,
} satisfies Record<keyof CommentRangeEnd, ReservedValueDisposition>;

export type ExhaustiveCommentRangeEndReserved = ExhaustiveFields<
  CommentRangeEnd,
  keyof typeof COMMENT_RANGE_END_RESERVED
>;

export const COMMENT_REFERENCE_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
} satisfies Record<keyof CommentReference, ReservedValueDisposition>;

export type ExhaustiveCommentReferenceReserved = ExhaustiveFields<
  CommentReference,
  keyof typeof COMMENT_REFERENCE_RESERVED
>;

export const MATH_EQUATION_RESERVED = {
  type: NO_RESERVED_VALUE,
  display: NO_RESERVED_VALUE,
  ommlXml: NO_RESERVED_VALUE,
  plainText: NO_RESERVED_VALUE,
} satisfies Record<keyof MathEquation, ReservedValueDisposition>;

export type ExhaustiveMathEquationReserved = ExhaustiveFields<
  MathEquation,
  keyof typeof MATH_EQUATION_RESERVED
>;

export const TRACKED_CHANGE_INFO_RESERVED = {
  id: NO_RESERVED_VALUE,
  author: NO_RESERVED_VALUE,
  date: NO_RESERVED_VALUE,
  initials: NO_RESERVED_VALUE,
  utcDate: NO_RESERVED_VALUE,
} satisfies Record<keyof TrackedChangeInfo, ReservedValueDisposition>;

export type ExhaustiveTrackedChangeInfoReserved = ExhaustiveFields<
  TrackedChangeInfo,
  keyof typeof TRACKED_CHANGE_INFO_RESERVED
>;

export const TRACKED_CHANGE_UTC_DATE_RESERVED = {
  attribute: NO_RESERVED_VALUE,
  value: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<TrackedChangeInfo["utcDate"]>, ReservedValueDisposition>;

export type ExhaustiveTrackedChangeUtcDateReserved = ExhaustiveFields<
  NonNullable<TrackedChangeInfo["utcDate"]>,
  keyof typeof TRACKED_CHANGE_UTC_DATE_RESERVED
>;

export const PROPERTY_CHANGE_INFO_RESERVED = {
  rsid: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  author: NO_RESERVED_VALUE,
  date: NO_RESERVED_VALUE,
  initials: NO_RESERVED_VALUE,
  utcDate: NO_RESERVED_VALUE,
} satisfies Record<keyof PropertyChangeInfo, ReservedValueDisposition>;

export type ExhaustivePropertyChangeInfoReserved = ExhaustiveFields<
  PropertyChangeInfo,
  keyof typeof PROPERTY_CHANGE_INFO_RESERVED
>;

export const INSERTION_RESERVED = {
  type: NO_RESERVED_VALUE,
  info: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
} satisfies Record<keyof Insertion, ReservedValueDisposition>;

export type ExhaustiveInsertionReserved = ExhaustiveFields<
  Insertion,
  keyof typeof INSERTION_RESERVED
>;

export const DELETION_RESERVED = {
  type: NO_RESERVED_VALUE,
  info: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
} satisfies Record<keyof Deletion, ReservedValueDisposition>;

export type ExhaustiveDeletionReserved = ExhaustiveFields<Deletion, keyof typeof DELETION_RESERVED>;

export const MOVE_FROM_RESERVED = {
  type: NO_RESERVED_VALUE,
  info: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
} satisfies Record<keyof MoveFrom, ReservedValueDisposition>;

export type ExhaustiveMoveFromReserved = ExhaustiveFields<
  MoveFrom,
  keyof typeof MOVE_FROM_RESERVED
>;

export const MOVE_TO_RESERVED = {
  type: NO_RESERVED_VALUE,
  info: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
} satisfies Record<keyof MoveTo, ReservedValueDisposition>;

export type ExhaustiveMoveToReserved = ExhaustiveFields<MoveTo, keyof typeof MOVE_TO_RESERVED>;

export const MOVE_FROM_RANGE_START_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  name: NO_RESERVED_VALUE,
  author: NO_RESERVED_VALUE,
  colFirst: NO_RESERVED_VALUE,
  colLast: NO_RESERVED_VALUE,
  date: NO_RESERVED_VALUE,
  displacedByCustomXml: NO_RESERVED_VALUE,
} satisfies Record<keyof MoveFromRangeStart, ReservedValueDisposition>;

export type ExhaustiveMoveFromRangeStartReserved = ExhaustiveFields<
  MoveFromRangeStart,
  keyof typeof MOVE_FROM_RANGE_START_RESERVED
>;

export const MOVE_FROM_RANGE_END_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  displacedByCustomXml: NO_RESERVED_VALUE,
} satisfies Record<keyof MoveFromRangeEnd, ReservedValueDisposition>;

export type ExhaustiveMoveFromRangeEndReserved = ExhaustiveFields<
  MoveFromRangeEnd,
  keyof typeof MOVE_FROM_RANGE_END_RESERVED
>;

export const MOVE_TO_RANGE_START_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  name: NO_RESERVED_VALUE,
  author: NO_RESERVED_VALUE,
  colFirst: NO_RESERVED_VALUE,
  colLast: NO_RESERVED_VALUE,
  date: NO_RESERVED_VALUE,
  displacedByCustomXml: NO_RESERVED_VALUE,
} satisfies Record<keyof MoveToRangeStart, ReservedValueDisposition>;

export type ExhaustiveMoveToRangeStartReserved = ExhaustiveFields<
  MoveToRangeStart,
  keyof typeof MOVE_TO_RANGE_START_RESERVED
>;

export const MOVE_TO_RANGE_END_RESERVED = {
  type: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  displacedByCustomXml: NO_RESERVED_VALUE,
} satisfies Record<keyof MoveToRangeEnd, ReservedValueDisposition>;

export type ExhaustiveMoveToRangeEndReserved = ExhaustiveFields<
  MoveToRangeEnd,
  keyof typeof MOVE_TO_RANGE_END_RESERVED
>;

export const RUN_PROPERTY_CHANGE_RESERVED = {
  type: NO_RESERVED_VALUE,
  info: NO_RESERVED_VALUE,
  previousFormatting: NO_RESERVED_VALUE,
  currentFormatting: NO_RESERVED_VALUE,
} satisfies Record<keyof RunPropertyChange, ReservedValueDisposition>;

export type ExhaustiveRunPropertyChangeReserved = ExhaustiveFields<
  RunPropertyChange,
  keyof typeof RUN_PROPERTY_CHANGE_RESERVED
>;

export const PARAGRAPH_PROPERTY_CHANGE_RESERVED = {
  type: NO_RESERVED_VALUE,
  info: NO_RESERVED_VALUE,
  previousFormatting: NO_RESERVED_VALUE,
  currentFormatting: NO_RESERVED_VALUE,
} satisfies Record<keyof ParagraphPropertyChange, ReservedValueDisposition>;

export type ExhaustiveParagraphPropertyChangeReserved = ExhaustiveFields<
  ParagraphPropertyChange,
  keyof typeof PARAGRAPH_PROPERTY_CHANGE_RESERVED
>;

export const TABLE_PROPERTY_CHANGE_RESERVED = {
  type: NO_RESERVED_VALUE,
  info: NO_RESERVED_VALUE,
  previousFormatting: NO_RESERVED_VALUE,
  currentFormatting: NO_RESERVED_VALUE,
} satisfies Record<keyof TablePropertyChange, ReservedValueDisposition>;

export type ExhaustiveTablePropertyChangeReserved = ExhaustiveFields<
  TablePropertyChange,
  keyof typeof TABLE_PROPERTY_CHANGE_RESERVED
>;

export const TABLE_ROW_PROPERTY_CHANGE_RESERVED = {
  type: NO_RESERVED_VALUE,
  info: NO_RESERVED_VALUE,
  previousFormatting: NO_RESERVED_VALUE,
  currentFormatting: NO_RESERVED_VALUE,
} satisfies Record<keyof TableRowPropertyChange, ReservedValueDisposition>;

export type ExhaustiveTableRowPropertyChangeReserved = ExhaustiveFields<
  TableRowPropertyChange,
  keyof typeof TABLE_ROW_PROPERTY_CHANGE_RESERVED
>;

export const TABLE_CELL_PROPERTY_CHANGE_RESERVED = {
  type: NO_RESERVED_VALUE,
  info: NO_RESERVED_VALUE,
  previousFormatting: NO_RESERVED_VALUE,
  currentFormatting: NO_RESERVED_VALUE,
} satisfies Record<keyof TableCellPropertyChange, ReservedValueDisposition>;

export type ExhaustiveTableCellPropertyChangeReserved = ExhaustiveFields<
  TableCellPropertyChange,
  keyof typeof TABLE_CELL_PROPERTY_CHANGE_RESERVED
>;

export const SECTION_PROPERTY_CHANGE_RESERVED = {
  type: NO_RESERVED_VALUE,
  info: NO_RESERVED_VALUE,
  previousProperties: NO_RESERVED_VALUE,
  previousReferences: NO_RESERVED_VALUE,
  currentProperties: NO_RESERVED_VALUE,
} satisfies Record<keyof SectionPropertyChange, ReservedValueDisposition>;

export type ExhaustiveSectionPropertyChangeReserved = ExhaustiveFields<
  SectionPropertyChange,
  keyof typeof SECTION_PROPERTY_CHANGE_RESERVED
>;

export const SECTION_PROPERTY_CHANGE_REFERENCES_RESERVED = {
  headerReferences: NO_RESERVED_VALUE,
  footerReferences: NO_RESERVED_VALUE,
} satisfies Record<
  keyof NonNullable<SectionPropertyChange["previousReferences"]>,
  ReservedValueDisposition
>;

export type ExhaustiveSectionPropertyChangeReferencesReserved = ExhaustiveFields<
  NonNullable<SectionPropertyChange["previousReferences"]>,
  keyof typeof SECTION_PROPERTY_CHANGE_REFERENCES_RESERVED
>;

export const TABLE_STRUCTURAL_CHANGE_INFO_RESERVED = {
  type: NO_RESERVED_VALUE,
  info: NO_RESERVED_VALUE,
  verticalMerge: readerOwned({
    slot: "w:vMerge@val",
    sentinel: "absent|continue",
    reader: RESERVED_VALUE_READERS.tableCellProperties,
    evidence: "vmerge-absent-means-continue",
  }),
  verticalMergeOriginal: readerOwned({
    slot: "w:vMerge@val",
    sentinel: "absent|continue",
    reader: RESERVED_VALUE_READERS.tableCellProperties,
    evidence: "vmerge-absent-means-continue",
  }),
} satisfies Record<UnionFields<TableStructuralChangeInfo>, ReservedValueDisposition>;

export const SDT_PROPERTIES_RESERVED = {
  sdtType: NO_RESERVED_VALUE,
  id: NO_RESERVED_VALUE,
  alias: NO_RESERVED_VALUE,
  tag: NO_RESERVED_VALUE,
  lock: NO_RESERVED_VALUE,
  placeholder: NO_RESERVED_VALUE,
  showingPlaceholder: toggle("w:showingPlcHdr@val"),
  dateFormat: NO_RESERVED_VALUE,
  dateValueISO: NO_RESERVED_VALUE,
  listItems: NO_RESERVED_VALUE,
  dropdownLastValue: notModelled({
    slot: "w:dropDownList@lastValue",
    sentinel: "",
    reason:
      "The XSD default is the empty string, so a dropdown that has never been used and one whose selection was cleared are the same bytes. folio stores the string as it finds it and never distinguishes the two.",
  }),
  checked: toggle("w:checked@val"),
  rawPropertiesXml: NO_RESERVED_VALUE,
  rawEndPropertiesXml: NO_RESERVED_VALUE,
  rawSdtChildrenBeforeContent: NO_RESERVED_VALUE,
  rawSdtChildrenAfterContent: NO_RESERVED_VALUE,
} satisfies Record<keyof SdtProperties, ReservedValueDisposition>;

export type ExhaustiveSdtPropertiesReserved = ExhaustiveFields<
  SdtProperties,
  keyof typeof SDT_PROPERTIES_RESERVED
>;

export const SDT_LIST_ITEM_RESERVED = {
  displayText: NO_RESERVED_VALUE,
  value: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<SdtProperties["listItems"]>[number], ReservedValueDisposition>;

export type ExhaustiveSdtListItemReserved = ExhaustiveFields<
  NonNullable<SdtProperties["listItems"]>[number],
  keyof typeof SDT_LIST_ITEM_RESERVED
>;

export const INLINE_SDT_RESERVED = {
  type: NO_RESERVED_VALUE,
  properties: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
} satisfies Record<keyof InlineSdt, ReservedValueDisposition>;

export type ExhaustiveInlineSdtReserved = ExhaustiveFields<
  InlineSdt,
  keyof typeof INLINE_SDT_RESERVED
>;

export const BLOCK_SDT_RESERVED = {
  type: NO_RESERVED_VALUE,
  properties: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
} satisfies Record<keyof BlockSdt, ReservedValueDisposition>;

export type ExhaustiveBlockSdtReserved = ExhaustiveFields<
  BlockSdt,
  keyof typeof BLOCK_SDT_RESERVED
>;

export const PARAGRAPH_MARK_CHANGE_RESERVED = {
  kind: NO_RESERVED_VALUE,
  info: NO_RESERVED_VALUE,
} satisfies Record<keyof ParagraphMarkChange, ReservedValueDisposition>;

export type ExhaustiveParagraphMarkChangeReserved = ExhaustiveFields<
  ParagraphMarkChange,
  keyof typeof PARAGRAPH_MARK_CHANGE_RESERVED
>;

export const PARAGRAPH_RESERVED = {
  type: NO_RESERVED_VALUE,
  paraId: NO_RESERVED_VALUE,
  textId: NO_RESERVED_VALUE,
  formatting: NO_RESERVED_VALUE,
  propertyChanges: NO_RESERVED_VALUE,
  pPrMark: NO_RESERVED_VALUE,
  reviewCarrier: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
  listRendering: NO_RESERVED_VALUE,
  renderedPageBreakBefore: NO_RESERVED_VALUE,
  sectionProperties: NO_RESERVED_VALUE,
  // Attributes replayed as the source wrote them. A reserved value is a
  // spelling the model interprets; this slot interprets nothing.
  preservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof Paragraph, ReservedValueDisposition>;

export type ExhaustiveParagraphReserved = ExhaustiveFields<
  Paragraph,
  keyof typeof PARAGRAPH_RESERVED
>;

export const HEADER_REFERENCE_RESERVED = {
  type: readerOwned({
    slot: "w:headerReference@type",
    sentinel: "default",
    reader: RESERVED_VALUE_READERS.headerFooterType,
  }),
  rId: NO_RESERVED_VALUE,
} satisfies Record<keyof HeaderReference, ReservedValueDisposition>;

export type ExhaustiveHeaderReferenceReserved = ExhaustiveFields<
  HeaderReference,
  keyof typeof HEADER_REFERENCE_RESERVED
>;

export const FOOTER_REFERENCE_RESERVED = {
  type: readerOwned({
    slot: "w:footerReference@type",
    sentinel: "default",
    reader: RESERVED_VALUE_READERS.headerFooterType,
  }),
  rId: NO_RESERVED_VALUE,
} satisfies Record<keyof FooterReference, ReservedValueDisposition>;

export type ExhaustiveFooterReferenceReserved = ExhaustiveFields<
  FooterReference,
  keyof typeof FOOTER_REFERENCE_RESERVED
>;

export const HEADER_FOOTER_RESERVED = {
  type: NO_RESERVED_VALUE,
  hdrFtrType: readerOwned({
    slot: "w:headerReference@type|w:footerReference@type",
    sentinel: "default",
    reader: RESERVED_VALUE_READERS.headerFooterType,
  }),
  content: NO_RESERVED_VALUE,
  watermark: NO_RESERVED_VALUE,
  rawWatermarkXml: NO_RESERVED_VALUE,
  verbatimXml: NO_RESERVED_VALUE,
  verbatimFingerprint: NO_RESERVED_VALUE,
  watermarkBlockIndex: NO_RESERVED_VALUE,
} satisfies Record<keyof HeaderFooter, ReservedValueDisposition>;

export type ExhaustiveHeaderFooterReserved = ExhaustiveFields<
  HeaderFooter,
  keyof typeof HEADER_FOOTER_RESERVED
>;

export const TEXT_WATERMARK_RESERVED = {
  kind: NO_RESERVED_VALUE,
  text: NO_RESERVED_VALUE,
  font: NO_RESERVED_VALUE,
  color: NO_RESERVED_VALUE,
  diagonal: NO_RESERVED_VALUE,
  opacity: NO_RESERVED_VALUE,
} satisfies Record<keyof TextWatermark, ReservedValueDisposition>;

export type ExhaustiveTextWatermarkReserved = ExhaustiveFields<
  TextWatermark,
  keyof typeof TEXT_WATERMARK_RESERVED
>;

export const PICTURE_WATERMARK_RESERVED = {
  kind: NO_RESERVED_VALUE,
  imageRId: NO_RESERVED_VALUE,
  imageTarget: NO_RESERVED_VALUE,
  imageTargetExternal: NO_RESERVED_VALUE,
  scale: NO_RESERVED_VALUE,
  widthPt: NO_RESERVED_VALUE,
  heightPt: NO_RESERVED_VALUE,
  washout: NO_RESERVED_VALUE,
} satisfies Record<keyof PictureWatermark, ReservedValueDisposition>;

export type ExhaustivePictureWatermarkReserved = ExhaustiveFields<
  PictureWatermark,
  keyof typeof PICTURE_WATERMARK_RESERVED
>;

export const SECTION_RESERVED = {
  properties: NO_RESERVED_VALUE,
  content: NO_RESERVED_VALUE,
  headers: NO_RESERVED_VALUE,
  footers: NO_RESERVED_VALUE,
} satisfies Record<keyof Section, ReservedValueDisposition>;

export type ExhaustiveSectionReserved = ExhaustiveFields<Section, keyof typeof SECTION_RESERVED>;

export const DOCUMENT_BODY_RESERVED = {
  content: NO_RESERVED_VALUE,
  sections: NO_RESERVED_VALUE,
  finalSectionProperties: NO_RESERVED_VALUE,
  comments: NO_RESERVED_VALUE,
} satisfies Record<keyof DocumentBody, ReservedValueDisposition>;

export type ExhaustiveDocumentBodyReserved = ExhaustiveFields<
  DocumentBody,
  keyof typeof DOCUMENT_BODY_RESERVED
>;
