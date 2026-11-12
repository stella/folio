// A field-totality gate in a test file is read by no compiler, so the
// paragraph formatting that states every field lives here, where the package
// typecheck compiles it.
import { paragraphNumberingReference } from "@stll/docx-core/model";

import type { ParagraphFormatting, RecordField } from "../../types/document";

/** A paragraph formatting that states every field. */
export const COMPLETE_PARAGRAPH_FORMATTING = {
  alignment: "center",
  bidi: false,
  kinsoku: true,
  autoSpaceDE: false,
  autoSpaceDN: true,
  overflowPunctuation: false,
  spaceBefore: 0,
  spaceAfter: 120,
  lineSpacing: 240,
  lineSpacingRule: "auto",
  snapToGrid: false,
  beforeAutospacing: true,
  afterAutospacing: false,
  spacingExplicit: { before: true, after: false },
  spacingPreservedAttributes: [
    {
      namespace: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      name: "beforeLines",
      value: "100",
    },
  ],
  indentLeft: 720,
  indentRight: 0,
  indentFirstLine: -360,
  hangingIndent: true,
  indentPreservedAttributes: [
    {
      namespace: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      name: "startChars",
      value: "200",
    },
  ],
  borders: { bottom: { style: "single", size: 8 } },
  shading: { fill: { rgb: "E0E0E0" }, pattern: "clear" },
  tabs: [{ position: 720, alignment: "left", leader: "dot" }],
  keepNext: true,
  keepLines: false,
  widowControl: true,
  pageBreakBefore: false,
  contextualSpacing: true,
  numPr: paragraphNumberingReference({ numId: 7, ilvl: 1 }),
  numPrFromStyle: paragraphNumberingReference({ numId: 7, ilvl: 1 }),
  numberingChangeXml:
    '<w:numberingChange w:id="1" w:author="Author" w:date="2026-01-01T00:00:00Z" w:original="%1."/>',
  numberingInsertionXml:
    '<w:ins w:id="2" w:author="Author" w:date="2026-01-01T00:00:00Z"/>',
  outlineLevel: { kind: "heading", level: 2 },
  styleId: "BodyText",
  frame: { dropCap: "drop", lines: 2 },
  suppressLineNumbers: false,
  suppressAutoHyphens: true,
  runProperties: { bold: true },
  runInWithNext: false,
  preserved: { children: [{ index: 32, xml: '<w:cnfStyle w:val="000000000000"/>' }] },
} as const satisfies {
  [Field in RecordField<ParagraphFormatting>]-?: ParagraphFormatting[Field];
};
