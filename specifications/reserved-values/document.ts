/** Reserved-value decisions for document settings and the package. */

import type { ExhaustiveFields } from "../../packages/docx-core/src/model/exhaustiveFields";
import type {
  Document,
  DocumentSettings,
  DocxPackage,
} from "../../packages/docx-core/src/model/document";
import { NO_RESERVED_VALUE, type ReservedValueDisposition, toggle } from "./disposition.ts";

export const DOCUMENT_SETTINGS_RESERVED = {
  compatibilityMode: NO_RESERVED_VALUE,
  defaultTabStop: NO_RESERVED_VALUE,
  evenAndOddHeaders: toggle("w:evenAndOddHeaders@val"),
  mirrorMargins: toggle("w:mirrorMargins@val"),
  adjustLineHeightInTable: toggle("w:adjustLineHeightInTable@val"),
  doNotUseIndentAsNumberingTabStop: toggle("w:doNotUseIndentAsNumberingTabStop@val"),
  updateFields: toggle("w:updateFields@val"),
  themeFontLang: NO_RESERVED_VALUE,
  autoHyphenation: toggle("w:autoHyphenation@val"),
  doNotHyphenateCaps: toggle("w:doNotHyphenateCaps@val"),
  consecutiveHyphenLimit: NO_RESERVED_VALUE,
  hyphenationZoneTwips: NO_RESERVED_VALUE,
  lineBreakRules: NO_RESERVED_VALUE,
  splitPageBreakAndParagraphMark: toggle("w:splitPgBreakAndParaMark@val"),
} satisfies Record<keyof DocumentSettings, ReservedValueDisposition>;

export type ExhaustiveDocumentSettingsReserved = ExhaustiveFields<
  DocumentSettings,
  keyof typeof DOCUMENT_SETTINGS_RESERVED
>;

export const THEME_FONT_LANG_RESERVED = {
  eastAsia: NO_RESERVED_VALUE,
  bidi: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<DocumentSettings["themeFontLang"]>, ReservedValueDisposition>;

export type ExhaustiveThemeFontLangReserved = ExhaustiveFields<
  NonNullable<DocumentSettings["themeFontLang"]>,
  keyof typeof THEME_FONT_LANG_RESERVED
>;

export const LINE_BREAK_RULES_RESERVED = {
  noLineBreaksBefore: NO_RESERVED_VALUE,
  noLineBreaksAfter: NO_RESERVED_VALUE,
  useLegacyEthiopicAmharicRules: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<DocumentSettings["lineBreakRules"]>, ReservedValueDisposition>;

export type ExhaustiveLineBreakRulesReserved = ExhaustiveFields<
  NonNullable<DocumentSettings["lineBreakRules"]>,
  keyof typeof LINE_BREAK_RULES_RESERVED
>;

/** `noLineBreaksBefore` and `noLineBreaksAfter` carry the same pair. */
type LineBreakRule = NonNullable<
  NonNullable<DocumentSettings["lineBreakRules"]>["noLineBreaksBefore"]
>;

export const LINE_BREAK_RULE_RESERVED = {
  language: NO_RESERVED_VALUE,
  characters: NO_RESERVED_VALUE,
} satisfies Record<keyof LineBreakRule, ReservedValueDisposition>;

export type ExhaustiveLineBreakRuleReserved = ExhaustiveFields<
  LineBreakRule,
  keyof typeof LINE_BREAK_RULE_RESERVED
>;

export const DOCX_PACKAGE_RESERVED = {
  conformanceClass: NO_RESERVED_VALUE,
  document: NO_RESERVED_VALUE,
  settings: NO_RESERVED_VALUE,
  styles: NO_RESERVED_VALUE,
  theme: NO_RESERVED_VALUE,
  numbering: NO_RESERVED_VALUE,
  fontTable: NO_RESERVED_VALUE,
  footnotes: NO_RESERVED_VALUE,
  endnotes: NO_RESERVED_VALUE,
  headers: NO_RESERVED_VALUE,
  footers: NO_RESERVED_VALUE,
  relationships: NO_RESERVED_VALUE,
  media: NO_RESERVED_VALUE,
  properties: NO_RESERVED_VALUE,
} satisfies Record<keyof DocxPackage, ReservedValueDisposition>;

export type ExhaustiveDocxPackageReserved = ExhaustiveFields<
  DocxPackage,
  keyof typeof DOCX_PACKAGE_RESERVED
>;

export const DOCX_PACKAGE_PROPERTIES_RESERVED = {
  title: NO_RESERVED_VALUE,
  subject: NO_RESERVED_VALUE,
  creator: NO_RESERVED_VALUE,
  keywords: NO_RESERVED_VALUE,
  description: NO_RESERVED_VALUE,
  lastModifiedBy: NO_RESERVED_VALUE,
  revision: NO_RESERVED_VALUE,
  created: NO_RESERVED_VALUE,
  modified: NO_RESERVED_VALUE,
} satisfies Record<keyof NonNullable<DocxPackage["properties"]>, ReservedValueDisposition>;

export type ExhaustiveDocxPackagePropertiesReserved = ExhaustiveFields<
  NonNullable<DocxPackage["properties"]>,
  keyof typeof DOCX_PACKAGE_PROPERTIES_RESERVED
>;

export const DOCUMENT_RESERVED = {
  package: NO_RESERVED_VALUE,
  originalBuffer: NO_RESERVED_VALUE,
  templateVariables: NO_RESERVED_VALUE,
  requiredFonts: NO_RESERVED_VALUE,
  warnings: NO_RESERVED_VALUE,
  parseWarnings: NO_RESERVED_VALUE,
} satisfies Record<keyof Document, ReservedValueDisposition>;

export type ExhaustiveDocumentReserved = ExhaustiveFields<Document, keyof typeof DOCUMENT_RESERVED>;
