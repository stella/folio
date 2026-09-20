/** Reserved-value decisions for the colour, border, and shading primitives. */

import type {
  BorderSpec,
  ColorValue,
  ShadingProperties,
} from "../../packages/docx-core/src/model/colors";
import type { ExhaustiveFields } from "../../packages/docx-core/src/model/exhaustiveFields";
import {
  NO_RESERVED_VALUE,
  notModelled,
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
  themeColor: notModelled({
    slot: "w:color@themeColor",
    sentinel: "none",
    reason:
      "ST_ThemeColor carries a `none` member that cancels an inherited theme slot. ThemeColorSlot models the 16 real slots only, so `none` parses as absent; the two differ where a style sets a theme colour a run must drop.",
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
    reader: RESERVED_VALUE_READERS.borderStyle,
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
