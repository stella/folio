/**
 * Type exports for @stll/folio
 *
 * Re-exports all public TypeScript types from the split type modules.
 */

export type {
  // Color & Styling Primitives
  ThemeColor,
  ThemeColorValue,
  SchemeColorSlot,
  ColorValue,
  BorderSpec,
  ShadingProperties,

  // Text Formatting
  UnderlineStyle,
  TextEffect,
  EmphasisMark,
  TextFormatting,

  // Paragraph Formatting
  TabStopAlignment,
  TabLeader,
  TabStop,
  LineSpacingRule,
  ParagraphAlignment,
  ParagraphFormatting,

  // Table Formatting
  TableWidthType,
  TableMeasurement,
  TableBorders,
  CellMargins,
  TableLook,
  FloatingTableProperties,
  TableGridChange,
  TableFormatting,
  TableRowFormatting,
  ConditionalFormatStyle,
  TableCellFormatting,

  // Run Content
  TextContent,
  PositionalTab,
  TabContent,
  BreakContent,
  SymbolContent,
  NoteReferenceContent,
  FieldCharContent,
  InstrTextContent,
  SoftHyphenContent,
  NoBreakHyphenContent,
  DrawingContent,
  DrawingRawXmlMode,
  ShapeContent,
  RunContent,
  Run,

  // Hyperlinks & Bookmarks
  Hyperlink,
  BookmarkStart,
  BookmarkEnd,

  // Fields
  FieldType,
  SimpleField,
  ComplexField,
  Field,

  // Images
  ImageSize,
  ImageWrap,
  ImagePosition,
  ImageTransform,
  ImagePadding,
  ImageCrop,
  Image,

  // Shapes & Text Boxes
  ShapeType,
  ShapeGeometryAdjustment,
  ShapeFill,
  ShapeOutline,
  ShapeTextBody,
  Shape,
  TextBox,

  // Tables
  TableCell,
  TableRow,
  Table,

  // Lists & Numbering
  NumberFormat,
  LevelLegacy,
  LevelOverride,
  LevelSuffix,
  ListLevel,
  AbstractNumbering,
  NumberingInstance,
  ListRendering,
  NumberingDefinitions,

  // Headers & Footers
  HeaderFooterType,
  HeaderReference,
  FooterReference,
  HeaderFooter,

  // Footnotes & Endnotes
  FootnotePosition,
  EndnotePosition,
  NoteNumberRestart,
  FootnoteProperties,
  EndnoteProperties,
  Footnote,
  Endnote,

  // Paragraph
  ParagraphContent,
  Paragraph,

  // Section Properties
  PageOrientation,
  SectionStart,
  VerticalAlign,
  LineNumberRestart,
  Column,
  SectionProperties,

  // Section & Document Body
  BlockContent,
  Section,
  DocumentBackground,
  DocumentBackgroundDrawing,
  DocumentBody,

  // Styles
  StyleType,
  Style,
  DocDefaults,
  StyleDefinitions,

  // Theme
  ThemeColorScheme,
  ThemeFont,
  ThemeFontScheme,
  Theme,

  // Font Table
  EmbeddedFontRef,
  FontCharset,
  FontInfo,
  FontTable,

  // Relationships
  RelationshipType,
  Relationship,
  RelationshipMap,

  // Media
  MediaFile,

  // Package & Document
  DocxConformanceClass,
  DocxPackage,
  Document,
} from "./document";
