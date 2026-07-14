import {
  DOCUMENT_PRESET_VERSION,
  DOCUMENT_STYLE_SET_VERSION,
  type DocumentPreset,
  type DocumentStyleSet,
} from "./types";

const CLAUSE_NUMBERING_ID = 1;
const DEFINITIONS_NUMBERING_ID = 2;
const RECITALS_NUMBERING_ID = 3;
const PARTIES_NUMBERING_ID = 4;
const BULLET_NUMBERING_ID = 5;

const CLAUSE_ABSTRACT_NUMBERING_ID = 1;
const DEFINITIONS_ABSTRACT_NUMBERING_ID = 2;
const PARENTHETICAL_ABSTRACT_NUMBERING_ID = 3;
const BULLET_ABSTRACT_NUMBERING_ID = 4;

export const STELLA_STYLE_SET_NAME = "Stella Style";

/**
 * A conservative legal drafting system inspired by common continental and
 * common-law contract structure. It uses explicit fonts and neutral punctuation
 * so the package does not depend on a locale-specific Office theme.
 */
export const createStellaStyleSet = (): DocumentStyleSet => ({
  version: DOCUMENT_STYLE_SET_VERSION,
  name: STELLA_STYLE_SET_NAME,
  initialParagraphStyleId: "BodyText",
  settings: { defaultTabStop: 720 },
  fontTable: {
    fonts: [
      { name: "Arial", family: "swiss", pitch: "variable" },
      { name: "Georgia", family: "roman", pitch: "variable" },
    ],
  },
  styles: {
    docDefaults: {
      rPr: {
        fontFamily: { ascii: "Arial", hAnsi: "Arial", cs: "Arial" },
        fontSize: 20,
        fontSizeCs: 20,
      },
    },
    styles: [
      {
        styleId: "Normal",
        type: "paragraph",
        name: "Normal",
        default: true,
        qFormat: true,
        uiPriority: 0,
        pPr: { alignment: "both", spaceAfter: 240 },
        rPr: {
          fontFamily: { ascii: "Arial", hAnsi: "Arial", cs: "Arial" },
          fontSize: 20,
          fontSizeCs: 20,
        },
      },
      {
        styleId: "DefaultParagraphFont",
        type: "character",
        name: "Default Paragraph Font",
        default: true,
        semiHidden: true,
        unhideWhenUsed: true,
      },
      {
        styleId: "BodyText",
        type: "paragraph",
        name: "Body Text",
        basedOn: "Normal",
        qFormat: true,
        uiPriority: 1,
      },
      {
        styleId: "Title",
        type: "paragraph",
        name: "Title",
        basedOn: "BodyText",
        next: "BodyText",
        qFormat: true,
        uiPriority: 2,
        pPr: { alignment: "center", keepNext: true, spaceAfter: 360 },
        rPr: { bold: true, fontSize: 32, fontSizeCs: 32 },
      },
      {
        styleId: "TitleWithSubtitle",
        type: "paragraph",
        name: "Title with Subtitle",
        basedOn: "Title",
        next: "Subtitle",
        qFormat: true,
        uiPriority: 3,
        pPr: { spaceAfter: 0 },
      },
      {
        styleId: "Subtitle",
        type: "paragraph",
        name: "Subtitle",
        basedOn: "Title",
        next: "BodyText",
        qFormat: true,
        uiPriority: 4,
        rPr: {
          fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
          italic: true,
          color: { rgb: "7F7F7F" },
        },
      },
      {
        styleId: "RecitalsHeading",
        type: "paragraph",
        name: "Recitals Heading",
        basedOn: "BodyText",
        next: "Recital",
        qFormat: true,
        uiPriority: 5,
        pPr: { keepNext: true },
        rPr: { bold: true },
      },
      {
        styleId: "Recital",
        type: "paragraph",
        name: "Recital",
        basedOn: "BodyText",
        qFormat: true,
        uiPriority: 6,
        pPr: { numPr: { numId: RECITALS_NUMBERING_ID, ilvl: 0 } },
      },
      {
        styleId: "AgreedTerms",
        type: "paragraph",
        name: "Agreed Terms",
        basedOn: "BodyText",
        next: "ClauseHeading1",
        qFormat: true,
        uiPriority: 7,
        pPr: { keepNext: true, spaceBefore: 240 },
        rPr: { bold: true },
      },
      ...createClauseStyles(),
      ...createDefinitionStyles(),
      {
        styleId: "Party",
        type: "paragraph",
        name: "Party",
        basedOn: "BodyText",
        qFormat: true,
        uiPriority: 30,
        pPr: { numPr: { numId: PARTIES_NUMBERING_ID, ilvl: 0 } },
      },
      {
        styleId: "ListParagraph",
        type: "paragraph",
        name: "List Paragraph",
        basedOn: "BodyText",
        qFormat: true,
        uiPriority: 31,
        pPr: { numPr: { numId: BULLET_NUMBERING_ID, ilvl: 0 } },
      },
      {
        styleId: "TableText",
        type: "paragraph",
        name: "Table Text",
        basedOn: "Normal",
        qFormat: true,
        uiPriority: 32,
        pPr: { spaceAfter: 0 },
      },
      {
        styleId: "ScheduleHeading",
        type: "paragraph",
        name: "Schedule Heading",
        basedOn: "Title",
        next: "BodyText",
        qFormat: true,
        uiPriority: 33,
        pPr: { pageBreakBefore: true, keepNext: true, spaceAfter: 240 },
      },
      {
        styleId: "FootnoteText",
        type: "paragraph",
        name: "Footnote Text",
        basedOn: "Normal",
        link: "FootnoteTextChar",
        semiHidden: true,
        unhideWhenUsed: true,
        uiPriority: 99,
        pPr: { spaceAfter: 0 },
        rPr: { fontSize: 18, fontSizeCs: 18 },
      },
      {
        styleId: "FootnoteTextChar",
        type: "character",
        name: "Footnote Text Char",
        basedOn: "DefaultParagraphFont",
        link: "FootnoteText",
        semiHidden: true,
        unhideWhenUsed: true,
        uiPriority: 99,
        rPr: { fontSize: 18, fontSizeCs: 18 },
      },
      {
        styleId: "FootnoteReference",
        type: "character",
        name: "Footnote Reference",
        basedOn: "DefaultParagraphFont",
        semiHidden: true,
        unhideWhenUsed: true,
        uiPriority: 99,
        rPr: { vertAlign: "superscript" },
      },
      {
        styleId: "Footer",
        type: "paragraph",
        name: "Footer",
        basedOn: "Normal",
        pPr: { alignment: "right", spaceAfter: 0 },
        rPr: { color: { rgb: "7F7F7F" }, fontSize: 18, fontSizeCs: 18 },
      },
      {
        styleId: "Hyperlink",
        type: "character",
        name: "Hyperlink",
        basedOn: "DefaultParagraphFont",
        unhideWhenUsed: true,
        uiPriority: 99,
        rPr: { color: { rgb: "0563C1" }, underline: { style: "single" } },
      },
      {
        styleId: "TableNormal",
        type: "table",
        name: "Normal Table",
        default: true,
        semiHidden: true,
        unhideWhenUsed: true,
        uiPriority: 99,
      },
      {
        styleId: "TableGrid",
        type: "table",
        name: "Table Grid",
        basedOn: "TableNormal",
        qFormat: true,
        uiPriority: 59,
        tblPr: {
          borders: {
            top: { style: "single", size: 4, color: { rgb: "BFBFBF" } },
            left: { style: "single", size: 4, color: { rgb: "BFBFBF" } },
            bottom: { style: "single", size: 4, color: { rgb: "BFBFBF" } },
            right: { style: "single", size: 4, color: { rgb: "BFBFBF" } },
            insideH: { style: "single", size: 4, color: { rgb: "BFBFBF" } },
            insideV: { style: "single", size: 4, color: { rgb: "BFBFBF" } },
          },
        },
      },
    ],
  },
  numbering: createLegalNumbering(),
});

export const createStellaStyleDocumentPreset = (): DocumentPreset => ({
  version: DOCUMENT_PRESET_VERSION,
  name: STELLA_STYLE_SET_NAME,
  styleSet: createStellaStyleSet(),
  sectionProperties: {
    pageWidth: 11_906,
    pageHeight: 16_838,
    orientation: "portrait",
    marginTop: 1440,
    marginBottom: 1440,
    marginLeft: 1440,
    marginRight: 1440,
    headerDistance: 708,
    footerDistance: 708,
    gutter: 0,
    columnCount: 1,
    columnSpace: 708,
    equalWidth: true,
    sectionStart: "nextPage",
    verticalAlign: "top",
  },
});

const createClauseStyles = (): DocumentStyleSet["styles"]["styles"] => {
  const levels = [
    {
      styleId: "ClauseHeading1",
      name: "Clause Heading",
      level: 0,
      bold: true,
      next: "ClauseParagraph1",
    },
    {
      styleId: "ClauseParagraph1",
      name: "Clause 1.1",
      level: 1,
      bold: false,
      next: "ClauseParagraph1",
    },
    {
      styleId: "ClauseParagraph2",
      name: "Clause (a)",
      level: 2,
      bold: false,
      next: "ClauseParagraph2",
    },
    {
      styleId: "ClauseParagraph3",
      name: "Clause (i)",
      level: 3,
      bold: false,
      next: "ClauseParagraph3",
    },
    {
      styleId: "ClauseParagraph4",
      name: "Clause (A)",
      level: 4,
      bold: false,
      next: "ClauseParagraph4",
    },
  ] as const;
  return levels.map(({ styleId, name, level, bold, next }) => {
    const style = {
      styleId,
      type: "paragraph",
      name,
      basedOn: "BodyText",
      next,
      qFormat: true,
      uiPriority: 10 + level,
      pPr: {
        numPr: { numId: CLAUSE_NUMBERING_ID, ilvl: level },
        keepNext: level < 2,
        spaceBefore: clauseSpaceBefore(level),
      },
    } satisfies DocumentStyleSet["styles"]["styles"][number];
    if (bold) {
      return Object.assign(style, { rPr: { bold: true } });
    }
    return style;
  });
};

const clauseSpaceBefore = (level: number): number => {
  if (level === 0) {
    return 360;
  }
  return 0;
};

const createDefinitionStyles = (): DocumentStyleSet["styles"]["styles"] => {
  const levels = [
    { styleId: "Definition1", name: "Definition (a)", level: 0 },
    { styleId: "Definition2", name: "Definition (i)", level: 1 },
    { styleId: "Definition3", name: "Definition (A)", level: 2 },
  ] as const;
  return levels.map(({ styleId, name, level }) => ({
    styleId,
    type: "paragraph",
    name,
    basedOn: "BodyText",
    qFormat: true,
    uiPriority: 20 + level,
    pPr: { numPr: { numId: DEFINITIONS_NUMBERING_ID, ilvl: level } },
  }));
};

const createLegalNumbering = (): NonNullable<DocumentStyleSet["numbering"]> => ({
  abstractNums: [
    {
      abstractNumId: CLAUSE_ABSTRACT_NUMBERING_ID,
      multiLevelType: "multilevel",
      levels: [
        legalLevel(0, "decimal", "%1", 567, 567, true),
        legalLevel(1, "decimal", "%1.%2", 567, 567, true),
        legalLevel(2, "lowerLetter", "(%3)", 1134, 567),
        legalLevel(3, "lowerRoman", "(%4)", 1701, 567),
        legalLevel(4, "upperLetter", "(%5)", 2268, 567),
      ],
    },
    {
      abstractNumId: DEFINITIONS_ABSTRACT_NUMBERING_ID,
      multiLevelType: "multilevel",
      levels: [
        legalLevel(0, "lowerLetter", "(%1)", 720, 360),
        legalLevel(1, "lowerRoman", "(%2)", 1440, 360),
        legalLevel(2, "upperLetter", "(%3)", 2160, 360),
      ],
    },
    {
      abstractNumId: PARENTHETICAL_ABSTRACT_NUMBERING_ID,
      multiLevelType: "singleLevel",
      levels: [legalLevel(0, "upperLetter", "(%1)", 720, 360)],
    },
    {
      abstractNumId: BULLET_ABSTRACT_NUMBERING_ID,
      multiLevelType: "singleLevel",
      levels: [legalLevel(0, "bullet", "•", 720, 360)],
    },
  ],
  nums: [
    { numId: CLAUSE_NUMBERING_ID, abstractNumId: CLAUSE_ABSTRACT_NUMBERING_ID },
    { numId: DEFINITIONS_NUMBERING_ID, abstractNumId: DEFINITIONS_ABSTRACT_NUMBERING_ID },
    { numId: RECITALS_NUMBERING_ID, abstractNumId: PARENTHETICAL_ABSTRACT_NUMBERING_ID },
    { numId: PARTIES_NUMBERING_ID, abstractNumId: PARENTHETICAL_ABSTRACT_NUMBERING_ID },
    { numId: BULLET_NUMBERING_ID, abstractNumId: BULLET_ABSTRACT_NUMBERING_ID },
  ],
});

type LegalNumberFormat = "bullet" | "decimal" | "lowerLetter" | "lowerRoman" | "upperLetter";

const legalLevel = (
  ilvl: number,
  numFmt: LegalNumberFormat,
  lvlText: string,
  indentLeft: number,
  hangingIndent: number,
  isLegal = false,
) => ({
  ilvl,
  start: 1,
  numFmt,
  lvlText,
  suffix: "tab" as const,
  ...(isLegal ? { isLgl: true } : {}),
  pPr: {
    indentLeft,
    indentFirstLine: hangingIndent,
    hangingIndent: true,
  },
});
