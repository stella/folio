/**
 * OOXML enum schemas for parser narrowing.
 *
 * Parsers throughout this directory read string-valued XML attributes
 * (`<w:u w:val="single"/>`, `<w:pgBorders w:val="dashed"/>`, …) that
 * the document model types as a small literal union. The historical
 * pattern was `getAttribute(...) as UnderlineStyle`, which lies to
 * the type system: a malformed file produces a typed value that
 * silently contains garbage and breaks downstream consumers.
 *
 * These valibot picklists are the single source of truth for each
 * OOXML enum. `narrowEnum(value, schema)` widens a parsed string to
 * the typed enum *iff* the value is recognised; unknown values
 * become `undefined` so callers can fall back to a default, matching
 * how Word degrades when it doesn't recognise a value.
 *
 * Picklist entries are checked at compile time against the TS union
 * via `satisfies` — a new variant added to the model type fails to
 * compile until the picklist is updated, giving the same compile-time
 * coverage gate as `switch-exhaustiveness-check`.
 */

import * as v from "valibot";

import {
  CONDITIONAL_STYLE_TYPE_VALUES,
  EMPHASIS_MARK_VALUES,
  FIELD_TYPE_VALUES,
  FLOATING_TABLE_X_SPEC_VALUES,
  FLOATING_TABLE_Y_SPEC_VALUES,
  FONT_HINT_VALUES,
  FONT_THEME_VALUES,
  FRAME_WRAP_VALUES,
  FRAME_X_ALIGN_VALUES,
  FRAME_Y_ALIGN_VALUES,
  HIGHLIGHT_COLOR_VALUES,
  IMAGE_HORIZONTAL_ALIGNMENT_VALUES,
  IMAGE_HORIZONTAL_RELATIVE_TO_VALUES,
  IMAGE_VERTICAL_ALIGNMENT_VALUES,
  IMAGE_VERTICAL_RELATIVE_TO_VALUES,
  IMAGE_WRAP_TEXT_VALUES,
  LEVEL_SUFFIX_VALUES,
  LINE_SPACING_RULE_VALUES,
  NUMBER_FORMAT_VALUES,
  PARAGRAPH_ALIGNMENT_VALUES,
  POSITIONAL_TAB_ALIGNMENT_VALUES,
  POSITIONAL_TAB_LEADER_VALUES,
  POSITIONAL_TAB_RELATIVE_TO_VALUES,
  SDT_LOCK_VALUES,
  SHADING_PATTERN_VALUES,
  SHAPE_TYPE_VALUES,
  STYLE_TYPE_VALUES,
  TAB_LEADER_VALUES,
  TAB_STOP_ALIGNMENT_VALUES,
  TABLE_JUSTIFICATION_VALUES,
  TABLE_ROW_HEIGHT_RULE_VALUES,
  TABLE_WIDTH_TYPE_VALUES,
  TEXT_DIRECTION_VALUES,
  TEXT_EFFECT_VALUES,
  UNDERLINE_STYLE_VALUES,
} from "../types/documentEnumValues";

/**
 * Narrow a raw attribute string to a picklist member, or `undefined`
 * if the value isn't recognised. The cast lives once inside the
 * valibot output type; callers receive a properly-typed value with
 * no `as` at the call site.
 */
export const narrowEnum = <T extends string>(
  value: string | null | undefined,
  schema: v.PicklistSchema<readonly T[], undefined>,
): T | undefined => {
  if (value === null || value === undefined) {
    return undefined;
  }
  const result = v.safeParse(schema, value);
  return result.success ? result.output : undefined;
};

// ---------------------------------------------------------------------------
// Run formatting enums
// ---------------------------------------------------------------------------

export const UnderlineStyleSchema: v.PicklistSchema<
  readonly [
    "none",
    "single",
    "words",
    "double",
    "thick",
    "dotted",
    "dottedHeavy",
    "dash",
    "dashedHeavy",
    "dashLong",
    "dashLongHeavy",
    "dotDash",
    "dashDotHeavy",
    "dotDotDash",
    "dashDotDotHeavy",
    "wave",
    "wavyHeavy",
    "wavyDouble",
  ],
  undefined
> = v.picklist(UNDERLINE_STYLE_VALUES);

export const HighlightColorSchema: v.PicklistSchema<
  readonly [
    "black",
    "blue",
    "cyan",
    "darkBlue",
    "darkCyan",
    "darkGray",
    "darkGreen",
    "darkMagenta",
    "darkRed",
    "darkYellow",
    "green",
    "lightGray",
    "magenta",
    "none",
    "red",
    "white",
    "yellow",
  ],
  undefined
> = v.picklist(HIGHLIGHT_COLOR_VALUES);

export const TextEffectSchema: v.PicklistSchema<
  readonly [
    "none",
    "blinkBackground",
    "lights",
    "antsBlack",
    "antsRed",
    "shimmer",
    "sparkle",
  ],
  undefined
> = v.picklist(TEXT_EFFECT_VALUES);

export const EmphasisMarkSchema: v.PicklistSchema<
  readonly ["none", "dot", "comma", "circle", "underDot"],
  undefined
> = v.picklist(EMPHASIS_MARK_VALUES);

export const FontThemeSchema: v.PicklistSchema<
  readonly [
    "majorAscii",
    "majorHAnsi",
    "majorEastAsia",
    "majorBidi",
    "minorAscii",
    "minorHAnsi",
    "minorEastAsia",
    "minorBidi",
  ],
  undefined
> = v.picklist(FONT_THEME_VALUES);

export const FontHintSchema: v.PicklistSchema<
  readonly ["default", "eastAsia", "cs"],
  undefined
> = v.picklist(FONT_HINT_VALUES);

// ---------------------------------------------------------------------------
// Paragraph formatting enums
// ---------------------------------------------------------------------------

export const ParagraphAlignmentSchema: v.PicklistSchema<
  readonly [
    "start",
    "center",
    "end",
    "both",
    "mediumKashida",
    "distribute",
    "numTab",
    "highKashida",
    "lowKashida",
    "thaiDistribute",
    "left",
    "right",
  ],
  undefined
> = v.picklist(PARAGRAPH_ALIGNMENT_VALUES);

export const LineSpacingRuleSchema: v.PicklistSchema<
  readonly ["auto", "exact", "atLeast"],
  undefined
> = v.picklist(LINE_SPACING_RULE_VALUES);

export const TabStopAlignmentSchema: v.PicklistSchema<
  readonly [
    "clear",
    "start",
    "center",
    "end",
    "decimal",
    "bar",
    "num",
    "left",
    "right",
  ],
  undefined
> = v.picklist(TAB_STOP_ALIGNMENT_VALUES);

export const TabLeaderSchema: v.PicklistSchema<
  readonly ["none", "dot", "hyphen", "underscore", "heavy", "middleDot"],
  undefined
> = v.picklist(TAB_LEADER_VALUES);

export const PositionalTabRelativeToSchema: v.PicklistSchema<
  readonly ["margin", "indent"],
  undefined
> = v.picklist(POSITIONAL_TAB_RELATIVE_TO_VALUES);

export const PositionalTabAlignmentSchema: v.PicklistSchema<
  readonly ["left", "center", "right"],
  undefined
> = v.picklist(POSITIONAL_TAB_ALIGNMENT_VALUES);

export const PositionalTabLeaderSchema: v.PicklistSchema<
  readonly ["none", "dot", "hyphen", "underscore", "middleDot"],
  undefined
> = v.picklist(POSITIONAL_TAB_LEADER_VALUES);

export const FrameXAlignSchema: v.PicklistSchema<
  readonly ["left", "center", "right", "inside", "outside"],
  undefined
> = v.picklist(FRAME_X_ALIGN_VALUES);

export const FrameYAlignSchema: v.PicklistSchema<
  readonly ["top", "center", "bottom", "inside", "outside", "inline"],
  undefined
> = v.picklist(FRAME_Y_ALIGN_VALUES);

export const FrameWrapSchema: v.PicklistSchema<
  readonly ["around", "auto", "none", "notBeside", "through", "tight"],
  undefined
> = v.picklist(FRAME_WRAP_VALUES);

// ---------------------------------------------------------------------------
// Structured document tag (SDT) enums
// ---------------------------------------------------------------------------

export const SdtLockSchema: v.PicklistSchema<
  readonly ["sdtLocked", "contentLocked", "sdtContentLocked", "unlocked"],
  undefined
> = v.picklist(SDT_LOCK_VALUES);

// ---------------------------------------------------------------------------
// Field types
// ---------------------------------------------------------------------------

export const FieldTypeSchema: v.PicklistSchema<
  readonly [
    "PAGE",
    "NUMPAGES",
    "NUMWORDS",
    "NUMCHARS",
    "DATE",
    "TIME",
    "CREATEDATE",
    "SAVEDATE",
    "PRINTDATE",
    "AUTHOR",
    "TITLE",
    "SUBJECT",
    "KEYWORDS",
    "COMMENTS",
    "FILENAME",
    "FILESIZE",
    "TEMPLATE",
    "DOCPROPERTY",
    "DOCVARIABLE",
    "REF",
    "PAGEREF",
    "NOTEREF",
    "HYPERLINK",
    "TOC",
    "TOA",
    "INDEX",
    "SEQ",
    "STYLEREF",
    "AUTONUM",
    "AUTONUMLGL",
    "AUTONUMOUT",
    "LISTNUM",
    "IF",
    "MERGEFIELD",
    "NEXT",
    "NEXTIF",
    "ASK",
    "SET",
    "QUOTE",
    "INCLUDETEXT",
    "INCLUDEPICTURE",
    "SYMBOL",
    "ADVANCE",
    "EDITTIME",
    "REVNUM",
    "SECTION",
    "SECTIONPAGES",
    "USERADDRESS",
    "USERNAME",
    "USERINITIALS",
    "UNKNOWN",
  ],
  undefined
> = v.picklist(FIELD_TYPE_VALUES);

// ---------------------------------------------------------------------------
// Style enums
// ---------------------------------------------------------------------------

export const StyleTypeSchema: v.PicklistSchema<
  readonly ["paragraph", "character", "numbering", "table"],
  undefined
> = v.picklist(STYLE_TYPE_VALUES);

export const ConditionalStyleTypeSchema: v.PicklistSchema<
  readonly [
    "band1Horz",
    "band1Vert",
    "band2Horz",
    "band2Vert",
    "firstCol",
    "firstRow",
    "lastCol",
    "lastRow",
    "neCell",
    "nwCell",
    "seCell",
    "swCell",
    "wholeTable",
  ],
  undefined
> = v.picklist(CONDITIONAL_STYLE_TYPE_VALUES);

// ---------------------------------------------------------------------------
// Table enums
// ---------------------------------------------------------------------------

export const TableAlignmentSchema: v.PicklistSchema<
  readonly ["center", "end", "left", "right", "start"],
  undefined
> = v.picklist(TABLE_JUSTIFICATION_VALUES);

export const TableWidthTypeSchema: v.PicklistSchema<
  readonly ["auto", "dxa", "nil", "pct"],
  undefined
> = v.picklist(TABLE_WIDTH_TYPE_VALUES);

export const TableRowHeightRuleSchema: v.PicklistSchema<
  readonly ["auto", "atLeast", "exact"],
  undefined
> = v.picklist(TABLE_ROW_HEIGHT_RULE_VALUES);

export const TextDirectionSchema: v.PicklistSchema<
  readonly [
    "tb",
    "rl",
    "lr",
    "tbV",
    "rlV",
    "lrV",
    "btLr",
    "lrTb",
    "lrTbV",
    "tbLrV",
    "tbRl",
    "tbRlV",
  ],
  undefined
> = v.picklist(TEXT_DIRECTION_VALUES);

// ---------------------------------------------------------------------------
// Shading enums
// ---------------------------------------------------------------------------

export const ShadingPatternSchema: v.PicklistSchema<
  readonly [
    "clear",
    "solid",
    "horzStripe",
    "vertStripe",
    "reverseDiagStripe",
    "diagStripe",
    "horzCross",
    "diagCross",
    "thinHorzStripe",
    "thinVertStripe",
    "thinReverseDiagStripe",
    "thinDiagStripe",
    "thinHorzCross",
    "thinDiagCross",
    "pct5",
    "pct10",
    "pct12",
    "pct15",
    "pct20",
    "pct25",
    "pct30",
    "pct35",
    "pct37",
    "pct40",
    "pct45",
    "pct50",
    "pct55",
    "pct60",
    "pct62",
    "pct65",
    "pct70",
    "pct75",
    "pct80",
    "pct85",
    "pct87",
    "pct90",
    "pct95",
    "nil",
  ],
  undefined
> = v.picklist(SHADING_PATTERN_VALUES);

// ---------------------------------------------------------------------------
// Floating table position enums
// ---------------------------------------------------------------------------

export const FloatingTableXSpecSchema: v.PicklistSchema<
  readonly ["left", "center", "right", "inside", "outside"],
  undefined
> = v.picklist(FLOATING_TABLE_X_SPEC_VALUES);

export const FloatingTableYSpecSchema: v.PicklistSchema<
  readonly ["top", "center", "bottom", "inside", "outside", "inline"],
  undefined
> = v.picklist(FLOATING_TABLE_Y_SPEC_VALUES);

// ---------------------------------------------------------------------------
// Image positioning / wrap enums
// ---------------------------------------------------------------------------

export const ImageHorizontalRelativeToSchema: v.PicklistSchema<
  readonly [
    "character",
    "column",
    "insideMargin",
    "leftMargin",
    "margin",
    "outsideMargin",
    "page",
    "rightMargin",
  ],
  undefined
> = v.picklist(IMAGE_HORIZONTAL_RELATIVE_TO_VALUES);

export const ImageHorizontalAlignmentSchema: v.PicklistSchema<
  readonly ["left", "right", "center", "inside", "outside"],
  undefined
> = v.picklist(IMAGE_HORIZONTAL_ALIGNMENT_VALUES);

export const ImageVerticalRelativeToSchema: v.PicklistSchema<
  readonly [
    "insideMargin",
    "line",
    "margin",
    "outsideMargin",
    "page",
    "paragraph",
    "topMargin",
    "bottomMargin",
  ],
  undefined
> = v.picklist(IMAGE_VERTICAL_RELATIVE_TO_VALUES);

export const ImageVerticalAlignmentSchema: v.PicklistSchema<
  readonly ["top", "bottom", "center", "inside", "outside"],
  undefined
> = v.picklist(IMAGE_VERTICAL_ALIGNMENT_VALUES);

export const ImageWrapTextSchema: v.PicklistSchema<
  readonly ["bothSides", "left", "right", "largest"],
  undefined
> = v.picklist(IMAGE_WRAP_TEXT_VALUES);

// ---------------------------------------------------------------------------
// Shape enums
// ---------------------------------------------------------------------------

export const ShapeTypeSchema: v.PicklistSchema<
  readonly [
    "rect",
    "roundRect",
    "ellipse",
    "triangle",
    "rtTriangle",
    "parallelogram",
    "trapezoid",
    "pentagon",
    "hexagon",
    "heptagon",
    "octagon",
    "decagon",
    "dodecagon",
    "star4",
    "star5",
    "star6",
    "star7",
    "star8",
    "star10",
    "star12",
    "star16",
    "star24",
    "star32",
    "line",
    "straightConnector1",
    "bentConnector2",
    "bentConnector3",
    "bentConnector4",
    "bentConnector5",
    "curvedConnector2",
    "curvedConnector3",
    "curvedConnector4",
    "curvedConnector5",
    "rightArrow",
    "leftArrow",
    "upArrow",
    "downArrow",
    "leftRightArrow",
    "upDownArrow",
    "quadArrow",
    "leftRightUpArrow",
    "bentArrow",
    "uturnArrow",
    "leftUpArrow",
    "bentUpArrow",
    "curvedRightArrow",
    "curvedLeftArrow",
    "curvedUpArrow",
    "curvedDownArrow",
    "stripedRightArrow",
    "notchedRightArrow",
    "homePlate",
    "chevron",
    "rightArrowCallout",
    "downArrowCallout",
    "leftArrowCallout",
    "upArrowCallout",
    "leftRightArrowCallout",
    "quadArrowCallout",
    "circularArrow",
    "flowChartProcess",
    "flowChartAlternateProcess",
    "flowChartDecision",
    "flowChartInputOutput",
    "flowChartPredefinedProcess",
    "flowChartInternalStorage",
    "flowChartDocument",
    "flowChartMultidocument",
    "flowChartTerminator",
    "flowChartPreparation",
    "flowChartManualInput",
    "flowChartManualOperation",
    "flowChartConnector",
    "flowChartOffpageConnector",
    "flowChartPunchedCard",
    "flowChartPunchedTape",
    "flowChartSummingJunction",
    "flowChartOr",
    "flowChartCollate",
    "flowChartSort",
    "flowChartExtract",
    "flowChartMerge",
    "flowChartOnlineStorage",
    "flowChartDelay",
    "flowChartMagneticTape",
    "flowChartMagneticDisk",
    "flowChartMagneticDrum",
    "flowChartDisplay",
    "wedgeRectCallout",
    "wedgeRoundRectCallout",
    "wedgeEllipseCallout",
    "cloudCallout",
    "borderCallout1",
    "borderCallout2",
    "borderCallout3",
    "accentCallout1",
    "accentCallout2",
    "accentCallout3",
    "callout1",
    "callout2",
    "callout3",
    "accentBorderCallout1",
    "accentBorderCallout2",
    "accentBorderCallout3",
    "actionButtonBlank",
    "actionButtonHome",
    "actionButtonHelp",
    "actionButtonInformation",
    "actionButtonBackPrevious",
    "actionButtonForwardNext",
    "actionButtonBeginning",
    "actionButtonEnd",
    "actionButtonReturn",
    "actionButtonDocument",
    "actionButtonSound",
    "actionButtonMovie",
    "irregularSeal1",
    "irregularSeal2",
    "frame",
    "halfFrame",
    "corner",
    "diagStripe",
    "chord",
    "arc",
    "bracketPair",
    "bracePair",
    "leftBracket",
    "rightBracket",
    "leftBrace",
    "rightBrace",
    "can",
    "cube",
    "bevel",
    "donut",
    "noSmoking",
    "blockArc",
    "foldedCorner",
    "smileyFace",
    "heart",
    "lightningBolt",
    "sun",
    "moon",
    "cloud",
    "snip1Rect",
    "snip2SameRect",
    "snip2DiagRect",
    "snipRoundRect",
    "round1Rect",
    "round2SameRect",
    "round2DiagRect",
    "plaque",
    "teardrop",
    "mathPlus",
    "mathMinus",
    "mathMultiply",
    "mathDivide",
    "mathEqual",
    "mathNotEqual",
    "gear6",
    "gear9",
    "funnel",
    "pieWedge",
    "pie",
    "leftCircularArrow",
    "leftRightCircularArrow",
    "swooshArrow",
    "textBox",
  ],
  undefined
> = v.picklist(SHAPE_TYPE_VALUES);

// ---------------------------------------------------------------------------
// Numbering enums
// ---------------------------------------------------------------------------

export const NumberFormatSchema: v.PicklistSchema<
  readonly [
    "decimal",
    "upperRoman",
    "lowerRoman",
    "upperLetter",
    "lowerLetter",
    "ordinal",
    "cardinalText",
    "ordinalText",
    "hex",
    "chicago",
    "ideographDigital",
    "japaneseCounting",
    "aiueo",
    "iroha",
    "decimalFullWidth",
    "decimalHalfWidth",
    "japaneseLegal",
    "japaneseDigitalTenThousand",
    "decimalEnclosedCircle",
    "decimalFullWidth2",
    "aiueoFullWidth",
    "irohaFullWidth",
    "decimalZero",
    "bullet",
    "ganada",
    "chosung",
    "decimalEnclosedFullstop",
    "decimalEnclosedParen",
    "decimalEnclosedCircleChinese",
    "ideographEnclosedCircle",
    "ideographTraditional",
    "ideographZodiac",
    "ideographZodiacTraditional",
    "taiwaneseCounting",
    "ideographLegalTraditional",
    "taiwaneseCountingThousand",
    "taiwaneseDigital",
    "chineseCounting",
    "chineseLegalSimplified",
    "chineseCountingThousand",
    "koreanDigital",
    "koreanCounting",
    "koreanLegal",
    "koreanDigital2",
    "vietnameseCounting",
    "russianLower",
    "russianUpper",
    "none",
    "numberInDash",
    "hebrew1",
    "hebrew2",
    "arabicAlpha",
    "arabicAbjad",
    "hindiVowels",
    "hindiConsonants",
    "hindiNumbers",
    "hindiCounting",
    "thaiLetters",
    "thaiNumbers",
    "thaiCounting",
    "bahtText",
    "dollarText",
    "custom",
  ],
  undefined
> = v.picklist(NUMBER_FORMAT_VALUES);

export const LevelSuffixSchema: v.PicklistSchema<
  readonly ["tab", "space", "nothing"],
  undefined
> = v.picklist(LEVEL_SUFFIX_VALUES);
