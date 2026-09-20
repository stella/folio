/** Reserved-value decisions for the colour, border, and shading primitives. */

import type {
  BorderSpec,
  ColorValue,
  ShadingProperties,
} from "../../packages/docx-core/src/model/colors";
import type { ExhaustiveFields } from "../../packages/docx-core/src/model/exhaustiveFields";
import {
  NO_RESERVED_VALUE,
  readerOwned,
  type ReservedValueDisposition,
  toggle,
} from "./disposition";
import { RESERVED_VALUE_READERS } from "./readers";

/**
 * `ST_HexColor` is a union of `ST_HexColorRGB` and the single-member
 * `ST_HexColorAuto`, so every slot typed with it accepts `auto`. These are the
 * ones folio parses into a {@link ColorValue}; the rest reach the model only
 * inside verbatim-captured markup.
 */
/**
 * Every slot typed `ST_ThemeColor`: a theme reference on a colour, a border
 * side, an underline, a page background, and both `w:shd` colours.
 */
const THEME_COLOR_SLOTS =
  "w:color@themeColor|w:shd@themeColor|w:shd@themeFill|w:u@themeColor|w:bdr@themeColor|w:background@themeColor|w:top@themeColor|w:bottom@themeColor|w:left@themeColor|w:right@themeColor|w:between@themeColor|w:bar@themeColor|w:insideH@themeColor|w:insideV@themeColor|w:tl2br@themeColor|w:tr2bl@themeColor";

const HEX_COLOR_SLOTS =
  "w:color@val|w:shd@fill|w:shd@color|w:u@color|w:bdr@color|w:background@color|w:top@color|w:bottom@color|w:left@color|w:right@color|w:between@color|w:bar@color|w:insideH@color|w:insideV@color|w:tl2br@color|w:tr2bl@color";

export const COLOR_VALUE_RESERVED = {
  // `auto` reaches the model twice: parsers that keep the sentinel set `auto`,
  // and parsers that keep the raw attribute leave the token in `rgb`.
  // `resolveColor` reads both spellings.
  rgb: readerOwned({
    slot: HEX_COLOR_SLOTS,
    sentinel: "auto",
    reader: RESERVED_VALUE_READERS.color,
    evidence: "hex-color-auto-is-context-dependent",
  }),
  // `ST_ThemeColor` is generated from the schema, so the union is the
  // enumeration; `themeColorSlot` is the one reader that turns a member into a
  // theme slot, and the one place `none` means "paints nothing".
  themeColor: readerOwned({
    slot: THEME_COLOR_SLOTS,
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.themeColor,
  }),
  themeTint: NO_RESERVED_VALUE,
  themeShade: NO_RESERVED_VALUE,
  auto: readerOwned({
    slot: HEX_COLOR_SLOTS,
    sentinel: "auto",
    reader: RESERVED_VALUE_READERS.color,
    evidence: "hex-color-auto-is-context-dependent",
  }),
} satisfies Record<keyof ColorValue, ReservedValueDisposition>;

export type ExhaustiveColorValueReserved = ExhaustiveFields<
  ColorValue,
  keyof typeof COLOR_VALUE_RESERVED
>;

/** Every `CT_Border` slot folio parses; `w:pgBorders` reuses the side names. */
const BORDER_STYLE_SLOTS =
  "w:top@val|w:bottom@val|w:left@val|w:right@val|w:between@val|w:bar@val|w:insideH@val|w:insideV@val|w:tl2br@val|w:tr2bl@val|w:bdr@val";

export const BORDER_SPEC_RESERVED = {
  style: readerOwned({
    slot: BORDER_STYLE_SLOTS,
    sentinel: "nil|none",
    reader: RESERVED_VALUE_READERS.borderSpec,
    evidence: "border-nil-and-none-are-distinct",
  }),
  color: NO_RESERVED_VALUE,
  size: NO_RESERVED_VALUE,
  space: NO_RESERVED_VALUE,
  shadow: toggle("w:bdr@shadow"),
  frame: toggle("w:bdr@frame"),
  artRelationshipId: NO_RESERVED_VALUE,
  topLeftArtRelationshipId: NO_RESERVED_VALUE,
  topRightArtRelationshipId: NO_RESERVED_VALUE,
  bottomLeftArtRelationshipId: NO_RESERVED_VALUE,
  bottomRightArtRelationshipId: NO_RESERVED_VALUE,
} satisfies Record<keyof BorderSpec, ReservedValueDisposition>;

export type ExhaustiveBorderSpecReserved = ExhaustiveFields<
  BorderSpec,
  keyof typeof BORDER_SPEC_RESERVED
>;

export const SHADING_PROPERTIES_RESERVED = {
  color: NO_RESERVED_VALUE,
  fill: NO_RESERVED_VALUE,
  pattern: readerOwned({
    slot: "w:shd@val",
    sentinel: "nil|clear",
    reader: RESERVED_VALUE_READERS.shading,
  }),
} satisfies Record<keyof ShadingProperties, ReservedValueDisposition>;

export type ExhaustiveShadingPropertiesReserved = ExhaustiveFields<
  ShadingProperties,
  keyof typeof SHADING_PROPERTIES_RESERVED
>;
