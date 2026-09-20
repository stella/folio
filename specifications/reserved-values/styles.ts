/** Reserved-value decisions for styles, theme, fonts, and package resources. */

import type { ExhaustiveFields } from "../../packages/docx-core/src/model/exhaustiveFields";
import type {
  DocDefaults,
  FontInfo,
  FontTable,
  MediaFile,
  Relationship,
  Style,
  StyleDefinitions,
  Theme,
  ThemeColorScheme,
  ThemeFont,
  ThemeFontScheme,
} from "../../packages/docx-core/src/model/styles";
import {
  NO_RESERVED_VALUE,
  readerOwned,
  type ReservedValueDisposition,
  toggle,
} from "./disposition";
import { RESERVED_VALUE_READERS } from "./readers";

export const STYLE_RESERVED = {
  styleId: NO_RESERVED_VALUE,
  type: NO_RESERVED_VALUE,
  name: NO_RESERVED_VALUE,
  basedOn: readerOwned({
    slot: "w:basedOn@val",
    sentinel: "unresolvable-styleid",
    reader: RESERVED_VALUE_READERS.styleDefinitions,
    evidence: "unknown-styleid-falls-back-to-the-default-style",
  }),
  next: readerOwned({
    slot: "w:next@val",
    sentinel: "unresolvable-styleid",
    reader: RESERVED_VALUE_READERS.styleDefinitions,
    evidence: "unknown-styleid-falls-back-to-the-default-style",
  }),
  link: readerOwned({
    slot: "w:link@val",
    sentinel: "unresolvable-styleid",
    reader: RESERVED_VALUE_READERS.styleDefinitions,
    evidence: "unknown-styleid-falls-back-to-the-default-style",
  }),
  uiPriority: NO_RESERVED_VALUE,
  hidden: toggle("w:semiHidden@val"),
  semiHidden: toggle("w:semiHidden@val"),
  unhideWhenUsed: toggle("w:unhideWhenUsed@val"),
  qFormat: toggle("w:qFormat@val"),
  default: toggle("w:style@default"),
  personal: toggle("w:personal@val"),
  pPr: NO_RESERVED_VALUE,
  rPr: NO_RESERVED_VALUE,
  tblPr: NO_RESERVED_VALUE,
  trPr: NO_RESERVED_VALUE,
  tcPr: NO_RESERVED_VALUE,
  tblStylePr: NO_RESERVED_VALUE,
} satisfies Record<keyof Style, ReservedValueDisposition>;

export type ExhaustiveStyleReserved = ExhaustiveFields<Style, keyof typeof STYLE_RESERVED>;

export const STYLE_CONDITIONAL_RESERVED = {
  type: NO_RESERVED_VALUE,
  pPr: NO_RESERVED_VALUE,
  rPr: NO_RESERVED_VALUE,
  tblPr: NO_RESERVED_VALUE,
  trPr: NO_RESERVED_VALUE,
  tcPr: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<Style["tblStylePr"]>[number], ReservedValueDisposition>;

export type ExhaustiveStyleConditionalReserved = ExhaustiveFields<
  NonNullable<Style["tblStylePr"]>[number],
  keyof typeof STYLE_CONDITIONAL_RESERVED
>;

export const DOC_DEFAULTS_RESERVED = {
  rPr: NO_RESERVED_VALUE,
  pPr: NO_RESERVED_VALUE,
} satisfies Record<keyof DocDefaults, ReservedValueDisposition>;

export type ExhaustiveDocDefaultsReserved = ExhaustiveFields<
  DocDefaults,
  keyof typeof DOC_DEFAULTS_RESERVED
>;

export const STYLE_DEFINITIONS_RESERVED = {
  docDefaults: NO_RESERVED_VALUE,
  latentStyles: NO_RESERVED_VALUE,
  styles: NO_RESERVED_VALUE,
} satisfies Record<keyof StyleDefinitions, ReservedValueDisposition>;

export type ExhaustiveStyleDefinitionsReserved = ExhaustiveFields<
  StyleDefinitions,
  keyof typeof STYLE_DEFINITIONS_RESERVED
>;

export const LATENT_STYLES_RESERVED = {
  defLockedState: toggle("w:latentStyles@defLockedState"),
  defUIPriority: NO_RESERVED_VALUE,
  defSemiHidden: toggle("w:latentStyles@defSemiHidden"),
  defUnhideWhenUsed: toggle("w:latentStyles@defUnhideWhenUsed"),
  defQFormat: toggle("w:latentStyles@defQFormat"),
  count: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<StyleDefinitions["latentStyles"]>, ReservedValueDisposition>;

export type ExhaustiveLatentStylesReserved = ExhaustiveFields<
  NonNullable<StyleDefinitions["latentStyles"]>,
  keyof typeof LATENT_STYLES_RESERVED
>;

export const THEME_COLOR_SCHEME_RESERVED = {
  dk1: NO_RESERVED_VALUE,
  lt1: NO_RESERVED_VALUE,
  dk2: NO_RESERVED_VALUE,
  lt2: NO_RESERVED_VALUE,
  accent1: NO_RESERVED_VALUE,
  accent2: NO_RESERVED_VALUE,
  accent3: NO_RESERVED_VALUE,
  accent4: NO_RESERVED_VALUE,
  accent5: NO_RESERVED_VALUE,
  accent6: NO_RESERVED_VALUE,
  hlink: NO_RESERVED_VALUE,
  folHlink: NO_RESERVED_VALUE,
} satisfies Record<keyof ThemeColorScheme, ReservedValueDisposition>;

export type ExhaustiveThemeColorSchemeReserved = ExhaustiveFields<
  ThemeColorScheme,
  keyof typeof THEME_COLOR_SCHEME_RESERVED
>;

export const THEME_FONT_RESERVED = {
  latin: NO_RESERVED_VALUE,
  ea: NO_RESERVED_VALUE,
  cs: NO_RESERVED_VALUE,
  fonts: NO_RESERVED_VALUE,
} satisfies Record<keyof ThemeFont, ReservedValueDisposition>;

export type ExhaustiveThemeFontReserved = ExhaustiveFields<
  ThemeFont,
  keyof typeof THEME_FONT_RESERVED
>;

export const THEME_FONT_SCHEME_RESERVED = {
  majorFont: NO_RESERVED_VALUE,
  minorFont: NO_RESERVED_VALUE,
} satisfies Record<keyof ThemeFontScheme, ReservedValueDisposition>;

export type ExhaustiveThemeFontSchemeReserved = ExhaustiveFields<
  ThemeFontScheme,
  keyof typeof THEME_FONT_SCHEME_RESERVED
>;

export const THEME_RESERVED = {
  name: NO_RESERVED_VALUE,
  colorScheme: NO_RESERVED_VALUE,
  fontScheme: NO_RESERVED_VALUE,
  formatScheme: NO_RESERVED_VALUE,
} satisfies Record<keyof Theme, ReservedValueDisposition>;

export type ExhaustiveThemeReserved = ExhaustiveFields<Theme, keyof typeof THEME_RESERVED>;

export const THEME_FORMAT_SCHEME_RESERVED = {
  name: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<Theme["formatScheme"]>, ReservedValueDisposition>;

export type ExhaustiveThemeFormatSchemeReserved = ExhaustiveFields<
  NonNullable<Theme["formatScheme"]>,
  keyof typeof THEME_FORMAT_SCHEME_RESERVED
>;

export const FONT_INFO_RESERVED = {
  name: NO_RESERVED_VALUE,
  altName: NO_RESERVED_VALUE,
  panose1: NO_RESERVED_VALUE,
  charset: NO_RESERVED_VALUE,
  family: readerOwned({
    slot: "w:family@val",
    sentinel: "auto",
    reader: RESERVED_VALUE_READERS.fontTable,
  }),
  pitch: readerOwned({
    slot: "w:pitch@val",
    sentinel: "default",
    reader: RESERVED_VALUE_READERS.fontTable,
  }),
  sig: NO_RESERVED_VALUE,
  embedRegular: NO_RESERVED_VALUE,
  embedBold: NO_RESERVED_VALUE,
  embedItalic: NO_RESERVED_VALUE,
  embedBoldItalic: NO_RESERVED_VALUE,
  preserved: NO_RESERVED_VALUE,
  preservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof FontInfo, ReservedValueDisposition>;

export type ExhaustiveFontInfoReserved = ExhaustiveFields<
  FontInfo,
  keyof typeof FONT_INFO_RESERVED
>;

export const FONT_SIGNATURE_RESERVED = {
  usb0: NO_RESERVED_VALUE,
  usb1: NO_RESERVED_VALUE,
  usb2: NO_RESERVED_VALUE,
  usb3: NO_RESERVED_VALUE,
  csb0: NO_RESERVED_VALUE,
  csb1: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<FontInfo["sig"]>, ReservedValueDisposition>;

export type ExhaustiveFontSignatureReserved = ExhaustiveFields<
  NonNullable<FontInfo["sig"]>,
  keyof typeof FONT_SIGNATURE_RESERVED
>;

export const FONT_TABLE_RESERVED = {
  fonts: NO_RESERVED_VALUE,
  preserved: NO_RESERVED_VALUE,
  preservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof FontTable, ReservedValueDisposition>;

export type ExhaustiveFontTableReserved = ExhaustiveFields<
  FontTable,
  keyof typeof FONT_TABLE_RESERVED
>;

export const RELATIONSHIP_RESERVED = {
  id: NO_RESERVED_VALUE,
  type: NO_RESERVED_VALUE,
  target: NO_RESERVED_VALUE,
  targetMode: NO_RESERVED_VALUE,
} satisfies Record<keyof Relationship, ReservedValueDisposition>;

export type ExhaustiveRelationshipReserved = ExhaustiveFields<
  Relationship,
  keyof typeof RELATIONSHIP_RESERVED
>;

export const MEDIA_FILE_RESERVED = {
  path: NO_RESERVED_VALUE,
  filename: NO_RESERVED_VALUE,
  mimeType: NO_RESERVED_VALUE,
  data: NO_RESERVED_VALUE,
  base64: NO_RESERVED_VALUE,
  dataUrl: NO_RESERVED_VALUE,
} satisfies Record<keyof MediaFile, ReservedValueDisposition>;

export type ExhaustiveMediaFileReserved = ExhaustiveFields<
  MediaFile,
  keyof typeof MEDIA_FILE_RESERVED
>;
