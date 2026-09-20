/**
 * GENERATED FILE — do not edit.
 *
 * The model enumerations that mirror an OOXML simple type, member for
 * member, so a token the format declares cannot go unspelled.
 *
 * Derived from `specifications/generated/docx-transitional-schema.gen.json`.
 * Regenerate with:
 *
 *   bun run generate:ooxml-enumerations
 */

/**
 * `ST_Jc`: every token a `w:jc/@w:val` may carry.
 *
 * `start` and `end` are the direction-aware alignments, distinct from
 * `left` and `right`: in a right-to-left paragraph `start` sits at the
 * right margin. They are members in their own right, not spellings of
 * `left` and `right`, and the layout resolves them against the
 * paragraph's direction. `numTab` aligns to the list number's tab stop.
 */
export const PARAGRAPH_ALIGNMENTS = [
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
] as const;

export type ParagraphAlignment =
  | "start"
  | "center"
  | "end"
  | "both"
  | "mediumKashida"
  | "distribute"
  | "numTab"
  | "highKashida"
  | "lowKashida"
  | "thaiDistribute"
  | "left"
  | "right";

/**
 * `ST_TabJc`: every token a `w:tab/@w:val` may carry.
 *
 * `start` and `end` are the direction-aware members, the same distinction
 * `ST_Jc` draws; `clear` removes an inherited stop rather than declaring
 * one, and `num` is the stop a numbered paragraph's text hangs from.
 */
export const TAB_STOP_ALIGNMENTS = [
  "clear",
  "start",
  "center",
  "end",
  "decimal",
  "bar",
  "num",
  "left",
  "right",
] as const;

export type TabStopAlignment =
  | "clear"
  | "start"
  | "center"
  | "end"
  | "decimal"
  | "bar"
  | "num"
  | "left"
  | "right";

/**
 * `ST_TextDirection`: every token a `w:textDirection/@w:val` may carry, on
 * a table cell, a section or a paragraph.
 *
 * Twelve tokens for six flows: each flow has a short spelling and a long one
 * naming the character and line progressions in full.
 * {@link TEXT_DIRECTION_FLOWS} is the six, and
 * {@link TEXT_DIRECTION_FLOW_BY_TOKEN} pairs each token with its flow.
 */
export const TEXT_DIRECTIONS = [
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
] as const;

export type TextDirection =
  | "tb"
  | "rl"
  | "lr"
  | "tbV"
  | "rlV"
  | "lrV"
  | "btLr"
  | "lrTb"
  | "lrTbV"
  | "tbLrV"
  | "tbRl"
  | "tbRlV";

/**
 * The six text flows `ST_TextDirection` names, spelled the Strict way.
 *
 * Strict's `ST_TextDirection` (ECMA-376 Part 1 §17.18.93) enumerates exactly
 * these six; Transitional adds a second spelling of each (Part 4 §14.11.7).
 * Rendering is decided per flow, so a cell written `tbRl` and its Strict twin
 * written `rl` paint the same.
 */
export const TEXT_DIRECTION_FLOWS = [
  "tb",
  "rl",
  "lr",
  "tbV",
  "rlV",
  "lrV",
] as const;

export type TextDirectionFlow =
  | "tb"
  | "rl"
  | "lr"
  | "tbV"
  | "rlV"
  | "lrV";

/**
 * The flow each `ST_TextDirection` token names.
 *
 * ECMA-376 Part 4 §14.11.7 gives each Transitional-only token as semantically
 * equivalent to a Strict one: `btLr` to `lr`, `lrTb` to `tb`, `lrTbV`
 * to `tbV`, `tbLrV` to `lrV`, `tbRl` to `rl` and `tbRlV` to `rlV`.
 * A reader keeps the token as authored and a writer writes it back; only
 * rendering goes through the flow.
 */
export const TEXT_DIRECTION_FLOW_BY_TOKEN = {
  tb: "tb",
  rl: "rl",
  lr: "lr",
  tbV: "tbV",
  rlV: "rlV",
  lrV: "lrV",
  btLr: "lr",
  lrTb: "tb",
  lrTbV: "tbV",
  tbLrV: "lrV",
  tbRl: "rl",
  tbRlV: "rlV",
} as const satisfies Record<TextDirection, TextDirectionFlow>;

/**
 * `ST_NumberFormat`: every token a `w:numFmt/@w:val` may carry, on a
 * numbering level, a note's properties or a section's page numbers.
 *
 * `custom` counts by the token list in the sibling `@w:format` rather
 * than by a vocabulary of its own; `none` prints no counter at all.
 */
export const NUMBER_FORMATS = [
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
] as const;

export type NumberFormat =
  | "decimal"
  | "upperRoman"
  | "lowerRoman"
  | "upperLetter"
  | "lowerLetter"
  | "ordinal"
  | "cardinalText"
  | "ordinalText"
  | "hex"
  | "chicago"
  | "ideographDigital"
  | "japaneseCounting"
  | "aiueo"
  | "iroha"
  | "decimalFullWidth"
  | "decimalHalfWidth"
  | "japaneseLegal"
  | "japaneseDigitalTenThousand"
  | "decimalEnclosedCircle"
  | "decimalFullWidth2"
  | "aiueoFullWidth"
  | "irohaFullWidth"
  | "decimalZero"
  | "bullet"
  | "ganada"
  | "chosung"
  | "decimalEnclosedFullstop"
  | "decimalEnclosedParen"
  | "decimalEnclosedCircleChinese"
  | "ideographEnclosedCircle"
  | "ideographTraditional"
  | "ideographZodiac"
  | "ideographZodiacTraditional"
  | "taiwaneseCounting"
  | "ideographLegalTraditional"
  | "taiwaneseCountingThousand"
  | "taiwaneseDigital"
  | "chineseCounting"
  | "chineseLegalSimplified"
  | "chineseCountingThousand"
  | "koreanDigital"
  | "koreanCounting"
  | "koreanLegal"
  | "koreanDigital2"
  | "vietnameseCounting"
  | "russianLower"
  | "russianUpper"
  | "none"
  | "numberInDash"
  | "hebrew1"
  | "hebrew2"
  | "arabicAlpha"
  | "arabicAbjad"
  | "hindiVowels"
  | "hindiConsonants"
  | "hindiNumbers"
  | "hindiCounting"
  | "thaiLetters"
  | "thaiNumbers"
  | "thaiCounting"
  | "bahtText"
  | "dollarText"
  | "custom";
