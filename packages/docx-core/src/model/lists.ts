/**
 * Lists & Numbering Types
 *
 * Types for bullet lists, numbered lists, and numbering definitions.
 */

import type { TextFormatting, ParagraphFormatting } from "./formatting";
import type { PreservedAttribute, PreservedMarkup } from "./preservedMarkup";

// ============================================================================
// LISTS & NUMBERING
// ============================================================================

/**
 * Number format type
 */
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
  // Synthetic in-memory formats for Word's `w:numFmt w:val="custom"` with an
  // XSLT-style zero-padded format string ("001, 002, ...", "0001, ...",
  // "00001, ..."). Not OOXML enum values — never serialized (numbering.xml is
  // preserved as-is on save); they exist so the render pipeline can carry the
  // pad width through the existing NumberFormat plumbing.
  | "decimalZero3"
  | "decimalZero4"
  | "decimalZero5"
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
  | "dollarText";

/**
 * Multi-level suffix (what follows the number)
 */
export type LevelSuffix = "tab" | "space" | "nothing";

/**
 * `w:lvlJc`: how the marker sits in the space the level reserves for it.
 *
 * The full `ST_Jc` enumeration, because that is what `CT_Jc` declares and a
 * value the reader refuses takes the whole `w:lvlJc` with it. `start` and `end`
 * are the Strict logical-direction spellings of `left` and `right`; the kashida
 * and distribute members belong to justified body text and say nothing a
 * marker can be laid out by, which is why the marker alignment a level resolves
 * to is a narrower thing than the value it carries.
 */
export type LevelJustification =
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
 * `w:legacy`: the Word 6.0 numbering geometry a level asks for.
 *
 * `legacy` is the element's own `w:legacy` attribute rather than a `w:val`, so
 * an absent flag and an explicit `w:legacy="0"` are different states and the
 * element is recorded by presence.
 */
export type LevelLegacy = {
  legacy?: boolean;
  legacySpace?: number;
  legacyIndent?: number;
};

/**
 * List level definition
 */
export type ListLevel = {
  /** Level index (0-8) */
  ilvl: number;
  /** `w:tplc`: the template code Word keys the level's gallery entry by. */
  tplc?: string;
  /** `w:tentative`: the level is defined but no paragraph has used it yet. */
  tentative?: boolean;
  /** Starting number */
  start?: number;
  /** Number format */
  numFmt: NumberFormat;
  /** Level text (e.g., "%1." or "•") */
  lvlText: string;
  /** `w:lvlText@w:null`: the marker is a null character rather than the text. */
  lvlTextNull?: boolean;
  /** Justification */
  lvlJc?: LevelJustification;
  /** Suffix after number */
  suffix?: LevelSuffix;
  /** `w:pStyle`: the paragraph style this level numbers. */
  pStyle?: string;
  /** `w:lvlPicBulletId`: the `w:numPicBullet` this level draws its bullet from. */
  lvlPicBulletId?: number;
  /** Paragraph properties for this level */
  pPr?: ParagraphFormatting;
  /** Run properties for the number/bullet */
  rPr?: TextFormatting;
  /** Restart numbering from higher level */
  lvlRestart?: number;
  /** Is legal numbering style */
  isLgl?: boolean;
  /** Legacy settings */
  legacy?: LevelLegacy;
  /** Children of `w:lvl` the model has no field for, in source position. */
  preserved?: PreservedMarkup;
  /** Attributes of `w:lvl` the model has no field for. */
  preservedAttributes?: PreservedAttribute[];
};

/**
 * Abstract numbering definition (w:abstractNum)
 */
export type AbstractNumbering = {
  /** Abstract numbering ID */
  abstractNumId: number;
  /**
   * `w:nsid`: the identity Word keys this template by across documents.
   *
   * Together with `tmpl` it is how a list carried between documents is
   * recognised as the same list, so a rebuild that dropped it turned every
   * template into a new one.
   */
  nsid?: string;
  /** Multi-level type */
  multiLevelType?: "hybridMultilevel" | "multilevel" | "singleLevel";
  /** `w:tmpl`: the gallery template this definition was created from. */
  tmpl?: string;
  /** Numbering style link */
  numStyleLink?: string;
  /** Style link */
  styleLink?: string;
  /** Level definitions */
  levels: ListLevel[];
  /** Name */
  name?: string;
  /** Children of `w:abstractNum` the model has no field for, in source position. */
  preserved?: PreservedMarkup;
  /** Attributes of `w:abstractNum` the model has no field for. */
  preservedAttributes?: PreservedAttribute[];
};

/** One `w:lvlOverride`: what a concrete instance changes about one level. */
export type LevelOverride = {
  ilvl: number;
  startOverride?: number;
  lvl?: ListLevel;
  /** Children of `w:lvlOverride` the model has no field for, in source position. */
  preserved?: PreservedMarkup;
  /** Attributes of `w:lvlOverride` the model has no field for. */
  preservedAttributes?: PreservedAttribute[];
};

/**
 * Numbering instance (w:num)
 */
export type NumberingInstance = {
  /** Numbering ID (referenced by paragraphs) */
  numId: number;
  /** Reference to abstract numbering */
  abstractNumId: number;
  /** Level overrides */
  levelOverrides?: LevelOverride[];
  /** Children of `w:num` the model has no field for, in source position. */
  preserved?: PreservedMarkup;
  /** Attributes of `w:num` the model has no field for. */
  preservedAttributes?: PreservedAttribute[];
};

/** Typography from numbering-level `w:rPr` that applies to the marker glyphs. */
export type ListMarkerFormatting = Pick<
  TextFormatting,
  | "fontFamily"
  | "fontSize"
  | "fontSizeCs"
  | "bold"
  | "boldCs"
  | "italic"
  | "italicCs"
  | "rtl"
  | "cs"
>;

/**
 * Computed list rendering info
 */
export type ListRendering = {
  /** Computed marker text (e.g., "1.", "a)", "•") */
  marker: string;
  /** Source `w:lvlText` pattern retained so newly inserted siblings can be numbered. */
  markerTemplate?: string;
  /** List level (0-8) */
  level: number;
  /** Numbering ID */
  numId: number;
  /** Whether this is a bullet or numbered list */
  isBullet: boolean;
  /** Whether this level uses legal numbering (parent placeholders render decimal). */
  isLegal?: boolean;
  /** Number format type (decimal, lowerRoman, upperRoman, etc.) */
  numFmt?: NumberFormat;
  /** Whether the list marker is hidden (w:vanish on level rPr) */
  markerHidden?: boolean;
  /** Canonical numbering-level marker typography, in OOXML units. */
  markerFormatting?: ListMarkerFormatting;
  /** Horizontal alignment of the marker around the paragraph's list anchor. */
  markerAlignment?: "left" | "center" | "right";
  /**
   * `w:caps` on the numbering level rPr — the marker text renders in upper
   * case (e.g. "SCHEDULE 1" instead of "Schedule 1"). Apply at substitution
   * time so number-format letters (lowerLetter / lowerRoman) also flip.
   */
  markerAllCaps?: boolean;
  /**
   * `w:suff` (§17.9.25) — what follows the marker before body text.
   * `tab` (the OOXML default) grows the marker to the next tab stop; `space`
   * adds one space glyph; `nothing` lets body text butt against the marker.
   */
  markerSuffix?: LevelSuffix;
  /** Number format for each level from 0 through this paragraph's level. */
  levelNumFmts?: NumberFormat[];
  /**
   * `w:start` for each level from 0 through this paragraph's level. Layout
   * seeds each level's counter from it, so a list whose definition starts at
   * 5 renders "5., 6." rather than "1., 2.".
   */
  levelStarts?: number[];
  /** Abstract numbering definition shared by one or more numIds. */
  abstractNumId?: number;
  /** Start override for this numId/level, if the numbering instance defines one. */
  startOverride?: number;
  /**
   * Number of inline `LISTNUM` (default-list) fields the paragraph contains.
   * Each represents an implicit counter advance at `ilvl + 1` (Word's
   * default LISTNUM behaviour), so the next paragraph at that depth picks
   * up the next letter — e.g. an OutNum2 line carrying inline "(a)" must
   * be followed by an OutNum3 "(b)", not another "(a)".
   */
  implicitChildLevelAdvances?: number;
  /**
   * Column offset (in twips, from the marker zone's left edge) where the
   * second slot of a tab-separated marker should land. Set when LISTNUM is
   * folded into the marker so the cached "(a)" aligns vertically with the
   * deeper level's marker column. Uses the next ilvl's `hangingIndent`.
   */
  markerSecondSlotOffsetTwips?: number;
};

/**
 * Complete numbering definitions
 */
export type NumberingDefinitions = {
  /** Abstract numbering definitions */
  abstractNums: AbstractNumbering[];
  /** Numbering instances */
  nums: NumberingInstance[];
  /** Children of `w:numbering` the model has no field for, in source position. */
  preserved?: PreservedMarkup;
  /** Attributes of `w:numbering` the model has no field for. */
  preservedAttributes?: PreservedAttribute[];
};
