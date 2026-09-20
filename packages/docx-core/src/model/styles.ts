/**
 * Styles, Theme, Font Table, Relationships & Media Types
 *
 * Types for document-level definitions that don't form the content tree.
 */

import type {
  TextFormatting,
  ParagraphFormatting,
  TableFormatting,
  TableRowFormatting,
  TableCellFormatting,
} from "./formatting";
import type { PreservedAttribute, PreservedMarkup } from "./preservedMarkup";

// ============================================================================
// STYLES
// ============================================================================

/**
 * Style type
 */
export type StyleType = "paragraph" | "character" | "numbering" | "table";

/**
 * Style definition
 */
export type Style = {
  /** Style ID */
  styleId: string;
  /** Style type */
  type: StyleType;
  /** Display name */
  name?: string;
  /** Based on style ID */
  basedOn?: string;
  /** Next style after Enter (for paragraph styles) */
  next?: string;
  /** Linked style (paragraph/character pair) */
  link?: string;
  /** UI sort priority */
  uiPriority?: number;
  /** Hidden from UI */
  hidden?: boolean;
  /** Semi-hidden from UI */
  semiHidden?: boolean;
  /** Unhide when used */
  unhideWhenUsed?: boolean;
  /** Quick format in gallery */
  qFormat?: boolean;
  /** Is default style */
  default?: boolean;
  /** Personal style (custom) */
  personal?: boolean;
  /** Paragraph properties (for paragraph/table styles) */
  pPr?: ParagraphFormatting;
  /** Run properties */
  rPr?: TextFormatting;
  /** Table properties (for table styles) */
  tblPr?: TableFormatting;
  /** Table row properties */
  trPr?: TableRowFormatting;
  /** Table cell properties */
  tcPr?: TableCellFormatting;
  /** Conditional table style parts */
  tblStylePr?: {
    type:
      | "band1Horz"
      | "band1Vert"
      | "band2Horz"
      | "band2Vert"
      | "firstCol"
      | "firstRow"
      | "lastCol"
      | "lastRow"
      | "neCell"
      | "nwCell"
      | "seCell"
      | "swCell"
      | "wholeTable";
    pPr?: ParagraphFormatting;
    rPr?: TextFormatting;
    tblPr?: TableFormatting;
    trPr?: TableRowFormatting;
    tcPr?: TableCellFormatting;
  }[];
};

/**
 * Document defaults (w:docDefaults)
 */
export type DocDefaults = {
  /** Default run properties */
  rPr?: TextFormatting;
  /** Default paragraph properties */
  pPr?: ParagraphFormatting;
};

/**
 * Style definitions from styles.xml
 */
export type StyleDefinitions = {
  /** Document defaults */
  docDefaults?: DocDefaults;
  /** Latent styles */
  latentStyles?: {
    defLockedState?: boolean;
    defUIPriority?: number;
    defSemiHidden?: boolean;
    defUnhideWhenUsed?: boolean;
    defQFormat?: boolean;
    count?: number;
  };
  /** Style definitions */
  styles: Style[];
};

// ============================================================================
// THEME
// ============================================================================

/**
 * Theme color scheme (a:clrScheme)
 */
export type ThemeColorScheme = {
  /** Dark 1 color (usually black) */
  dk1?: string;
  /** Light 1 color (usually white) */
  lt1?: string;
  /** Dark 2 color */
  dk2?: string;
  /** Light 2 color */
  lt2?: string;
  /** Accent colors 1-6 */
  accent1?: string;
  accent2?: string;
  accent3?: string;
  accent4?: string;
  accent5?: string;
  accent6?: string;
  /** Hyperlink color */
  hlink?: string;
  /** Followed hyperlink color */
  folHlink?: string;
};

/**
 * Theme font (with script variants)
 */
export type ThemeFont = {
  /** Latin font */
  latin?: string;
  /** East Asian font */
  ea?: string;
  /** Complex script font */
  cs?: string;
  /** Script-specific fonts */
  fonts?: Record<string, string>;
};

/**
 * Theme font scheme (a:fontScheme)
 */
export type ThemeFontScheme = {
  /** Major font (headings) */
  majorFont?: ThemeFont;
  /** Minor font (body text) */
  minorFont?: ThemeFont;
};

/**
 * Theme (from theme1.xml)
 */
export type Theme = {
  /** Theme name */
  name?: string;
  /** Color scheme */
  colorScheme?: ThemeColorScheme;
  /** Font scheme */
  fontScheme?: ThemeFontScheme;
  /** Format scheme (fills, lines, effects) - simplified */
  formatScheme?: {
    name?: string;
  };
};

// ============================================================================
// FONT TABLE
// ============================================================================

/**
 * `w:charset`: the code page a font is encoded in, numbered and named.
 *
 * Two attributes of one element rather than two fields on the font, because
 * `w:charset` is the thing that carries both and a producer may write either.
 */
export type FontCharset = {
  /** `w:val`: the code page as a two-digit hexadecimal number. */
  val?: string;
  /** `w:characterSet`: the same set named (`ANSI_CHARSET`) rather than numbered. */
  characterSet?: string;
};

/**
 * One embedded face: the relationship to its binary and how to read it.
 *
 * The relationship id alone is not enough to use the face. `w:fontKey` is the
 * GUID the `.odttf` is obfuscated with, so a package that keeps the id and
 * drops the key keeps a pointer to bytes nothing can decode.
 */
export type EmbeddedFontRef = {
  /** `r:id` of the relationship to the font binary, in `word/_rels/fontTable.xml.rels`. */
  id: string;
  /** `w:fontKey`: the obfuscation GUID, e.g. `{XXXXXXXX-…}`. */
  fontKey?: string;
  /** `w:subsetted`: the embedded face carries only the glyphs the document uses. */
  subsetted?: boolean;
};

/**
 * Font info from fontTable.xml
 */
export type FontInfo = {
  /** Font name */
  name: string;
  /** Alternate names */
  altName?: string;
  /** Panose-1 classification */
  panose1?: string;
  /** Character set */
  charset?: FontCharset;
  /** Font family type */
  family?: "decorative" | "modern" | "roman" | "script" | "swiss" | "auto";
  /** Pitch (fixed or variable) */
  pitch?: "default" | "fixed" | "variable";
  /** Signature */
  sig?: {
    usb0?: string;
    usb1?: string;
    usb2?: string;
    usb3?: string;
    csb0?: string;
    csb1?: string;
  };
  /** Embedded font data reference */
  embedRegular?: EmbeddedFontRef;
  embedBold?: EmbeddedFontRef;
  embedItalic?: EmbeddedFontRef;
  embedBoldItalic?: EmbeddedFontRef;
  /** Children of `w:font` the model does not hold, in source position. */
  preserved?: PreservedMarkup;
  /** Attributes of `w:font` the model has no field for. */
  preservedAttributes?: PreservedAttribute[];
};

/**
 * Font table from fontTable.xml
 */
export type FontTable = {
  fonts: FontInfo[];
  /** Children of `w:fonts` that are not `w:font`, in source position. */
  preserved?: PreservedMarkup;
  /**
   * Attributes of `w:fonts` the model has no field for.
   *
   * Not the namespace declarations and not `mc:Ignorable`: a rebuilt part
   * binds the prefixes its own body uses and states which of them are
   * ignorable, so replaying the source's would name prefixes nothing binds.
   */
  preservedAttributes?: PreservedAttribute[];
};

// ============================================================================
// RELATIONSHIPS
// ============================================================================

/**
 * Relationship type
 */
export type RelationshipType = string;

/**
 * Relationship entry
 */
export type Relationship = {
  /** Relationship ID (e.g., "rId1") */
  id: string;
  /** Relationship type URI */
  type: RelationshipType;
  /** Target path or URL */
  target: string;
  /** Target mode */
  targetMode?: "External" | "Internal";
};

/**
 * Relationship map (keyed by rId)
 */
export type RelationshipMap = Map<string, Relationship>;

// ============================================================================
// MEDIA
// ============================================================================

/**
 * Media file from word/media/
 */
export type MediaFile = {
  /** File path in ZIP */
  path: string;
  /** Original filename */
  filename?: string;
  /** MIME type */
  mimeType: string;
  /** Binary data */
  data: ArrayBuffer;
  /** Base64 encoded data for rendering */
  base64?: string;
  /** Data URL for direct use in src attributes */
  dataUrl?: string;
};
