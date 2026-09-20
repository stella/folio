/**
 * Color & Styling Primitives
 *
 * Basic types used throughout OOXML for colors, borders, and shading.
 */

import type { BorderStyleValue } from "./borderStyle";
import type { PreservedAttribute } from "./preservedMarkup";
import type { ThemeColorValue } from "./themeColor";

/**
 * Color value - can be direct RGB, theme reference, or auto
 */
export type ColorValue = {
  /** RGB hex value without # (e.g., "FF0000") */
  rgb?: string;
  /**
   * `w:themeColor`/`w:themeFill` as written: an `ST_ThemeColor` token, or a
   * token outside it captured verbatim. `themeColorSlot` resolves either.
   */
  themeColor?: ThemeColorValue;
  /** Tint modifier (0-255 as hex string, e.g., "80") - makes color lighter */
  themeTint?: string;
  /** Shade modifier (0-255 as hex string) - makes color darker */
  themeShade?: string;
  /** Auto color - context-dependent (usually black for text) */
  auto?: boolean;
};

/**
 * Border specification for any border (paragraph, table, page)
 */
export type BorderSpec = {
  /**
   * `w:val`: an `ST_Border` member, or the raw token when the file wrote one
   * the schema does not declare. Ask `statesNoBorder`/`isBorderNone`/
   * `isBorderNil` rather than comparing it; `nil` and `none` are not synonyms.
   */
  style: BorderStyleValue;
  /** Color of the border */
  color?: ColorValue;
  /** Width in eighths of a point (1/8 pt) */
  size?: number;
  /** Spacing from text in points */
  space?: number;
  /** Shadow effect */
  shadow?: boolean;
  /** Frame effect */
  frame?: boolean;
  /**
   * Custom page-border art relationship id (`w:id` on `<w:pgBorders>` side
   * children). Preserved for round-trip; folio does not paint art glyphs.
   */
  artRelationshipId?: string;
  /** Custom page-border art relationship id for the top-left corner. */
  topLeftArtRelationshipId?: string;
  /** Custom page-border art relationship id for the top-right corner. */
  topRightArtRelationshipId?: string;
  /** Custom page-border art relationship id for the bottom-left corner. */
  bottomLeftArtRelationshipId?: string;
  /** Custom page-border art relationship id for the bottom-right corner. */
  bottomRightArtRelationshipId?: string;
  /**
   * `CT_Border` attributes this record has no field for.
   *
   * The remainder rides the record that holds the element's modelled fields,
   * the rule `attributeRemainder.ts` states for `w:p` and `w:tr`, applied one
   * level down: a border is an attribute bag, so an attribute folio does not
   * read is lost the moment any other attribute makes the element modelled.
   */
  preservedAttributes?: PreservedAttribute[];
};

/**
 * Shading/background properties
 */
export type ShadingProperties = {
  /** Pattern fill color */
  color?: ColorValue;
  /** Background fill color */
  fill?: ColorValue;
  /** Shading pattern type */
  pattern?:
    | "clear"
    | "solid"
    | "horzStripe"
    | "vertStripe"
    | "reverseDiagStripe"
    | "diagStripe"
    | "horzCross"
    | "diagCross"
    | "thinHorzStripe"
    | "thinVertStripe"
    | "thinReverseDiagStripe"
    | "thinDiagStripe"
    | "thinHorzCross"
    | "thinDiagCross"
    | "pct5"
    | "pct10"
    | "pct12"
    | "pct15"
    | "pct20"
    | "pct25"
    | "pct30"
    | "pct35"
    | "pct37"
    | "pct40"
    | "pct45"
    | "pct50"
    | "pct55"
    | "pct60"
    | "pct62"
    | "pct65"
    | "pct70"
    | "pct75"
    | "pct80"
    | "pct85"
    | "pct87"
    | "pct90"
    | "pct95"
    | "nil";
  /** `CT_Shd` attributes this record has no field for; see {@link BorderSpec}. */
  preservedAttributes?: PreservedAttribute[];
};
