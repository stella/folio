/**
 * The one reader for `w:shd`, shared by the run, paragraph, style and table
 * tiers.
 *
 * `w:color` and `w:fill` are `ST_HexColor`: either six hex digits or the
 * reserved token `auto`. `auto` is not a colour and paints nothing, but it is
 * distinct from an absent attribute: an explicit `w:fill="auto"` on a cell
 * cancels a fill the table style would otherwise supply, and dropping it also
 * loses the attribute the next time the paragraph is saved.
 */

import type { ShadingProperties } from "../types/document";
import { isValidHexColor } from "../utils/colorResolver";
import { readAttributeBag } from "./attributeRemainder";
import { narrowEnum, ShadingPatternSchema, ThemeColorSlotSchema } from "./parserEnums";
import { SHADING_ATTRIBUTES } from "./propertyElementAttributes";
import { getAttribute } from "./xmlParser";
import type { XmlElement } from "./xmlParser";

/** `ST_HexColor`'s reserved "let the consumer decide" token. */
const AUTOMATIC_COLOR = "auto";

const parseHexColor = (value: string | null): ShadingProperties["fill"] => {
  if (value === AUTOMATIC_COLOR) {
    return { auto: true };
  }
  if (value && isValidHexColor(value)) {
    return { rgb: value };
  }
  return undefined;
};

export function parseShading(shd: XmlElement | null): ShadingProperties | undefined {
  if (!shd) {
    return undefined;
  }

  const props: ShadingProperties = {};

  const color = parseHexColor(getAttribute(shd, "w", "color"));
  if (color) {
    props.color = color;
  }

  const fill = parseHexColor(getAttribute(shd, "w", "fill"));
  if (fill) {
    props.fill = fill;
  }

  const themeFill = narrowEnum(getAttribute(shd, "w", "themeFill"), ThemeColorSlotSchema);
  if (themeFill) {
    props.fill ??= {};
    props.fill.themeColor = themeFill;
  }

  const themeFillTint = getAttribute(shd, "w", "themeFillTint");
  if (themeFillTint && props.fill) {
    props.fill.themeTint = themeFillTint;
  }

  const themeFillShade = getAttribute(shd, "w", "themeFillShade");
  if (themeFillShade && props.fill) {
    props.fill.themeShade = themeFillShade;
  }

  const pattern = narrowEnum(getAttribute(shd, "w", "val"), ShadingPatternSchema);
  if (pattern) {
    props.pattern = pattern;
  }

  // Only when the element is modelled at all: a `w:shd` folio takes nothing
  // from is captured whole by its container's dispatcher, and a remainder as
  // well would write the same attributes twice.
  if (Object.keys(props).length === 0) {
    return undefined;
  }
  const preservedAttributes = readAttributeBag(shd, SHADING_ATTRIBUTES);
  if (preservedAttributes) {
    props.preservedAttributes = preservedAttributes;
  }

  return props;
}
