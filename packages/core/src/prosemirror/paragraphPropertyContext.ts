import type { AuthoredParagraphProperties, ParagraphMarkProperties } from "@stll/docx-core/model";
import type { ParagraphFormatting, TextFormatting } from "../types/document";

export const PARAGRAPH_SPACING_INHERITANCE_SOURCE = {
  documentDefault: "document-default",
  implicitDefaultStyle: "implicit-default-style",
  style: "style",
} as const;

export type ParagraphSpacingInheritanceSource =
  (typeof PARAGRAPH_SPACING_INHERITANCE_SOURCE)[keyof typeof PARAGRAPH_SPACING_INHERITANCE_SOURCE];

export type ParagraphSpacingInheritance = {
  before?: ParagraphSpacingInheritanceSource;
  after?: ParagraphSpacingInheritanceSource;
};

export type ParagraphMarkEffectiveProperties = {
  defaultTextFormatting?: TextFormatting;
  runInWithNext?: boolean;
};

export type ParagraphMarkProjectionContext = {
  authored: ParagraphMarkProperties;
  effective: ParagraphMarkEffectiveProperties;
};

export type SerializedParagraphPropertyProjectionContext = {
  inheritedPPr: AuthoredParagraphProperties;
  numberingLevelIndent: ParagraphFormatting["numberingLevelIndent"] | null;
  numPrFromStyle: ParagraphFormatting["numPr"] | null;
  paragraphMark: ParagraphMarkProjectionContext;
  spacingInheritance: ParagraphSpacingInheritance;
};
