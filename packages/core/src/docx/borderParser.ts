/**
 * The one reader for `CT_Border`, shared by the paragraph (`w:pBdr`), style,
 * table (`w:tblBorders`/`w:tcBorders`) and page (`w:pgBorders`) tiers.
 *
 * `w:val` is `ST_Border`, whose 193 members include two distinct "no border"
 * tokens: `nil` and `none`. They are not interchangeable downstream, and an
 * explicit one overrides a border inherited from the container, so the member
 * the author wrote is preserved exactly. A token outside the enumeration is
 * kept verbatim and reported, which is how the repo already treats `w:numFmt`,
 * `w:suff` and `w:tab`.
 */

import { borderStyleFrom, isBorderStyle, PARSE_WARNING_CODES } from "@stll/docx-core/model";

import type { BorderSpec, ColorValue } from "../types/document";
import { readAttributeBag } from "./attributeRemainder";
import type { ParseContext } from "./parseContext";
import { BORDER_ATTRIBUTES } from "./propertyElementAttributes";
import { parseThemeColorAttribute } from "./themeColorAttribute";
import { getAttribute, parseNumericAttribute, parseOnOffAttribute } from "./xmlParser";
import type { XmlElement } from "./xmlParser";

/** `ST_HexColor`'s reserved "let the consumer decide" token, as `shadingParser` spells it. */
const AUTOMATIC_COLOR = "auto";

const parseBorderColor = (border: XmlElement, context?: ParseContext): ColorValue | undefined => {
  const rgb = getAttribute(border, "w", "color");
  const themeColor = getAttribute(border, "w", "themeColor");
  const themeTint = getAttribute(border, "w", "themeTint");
  const themeShade = getAttribute(border, "w", "themeShade");
  if (rgb === null && themeColor === null && themeTint === null && themeShade === null) {
    return undefined;
  }

  const color: ColorValue = {};
  if (rgb === AUTOMATIC_COLOR) {
    color.auto = true;
  } else if (rgb) {
    color.rgb = rgb;
  }

  const parsedThemeColor = parseThemeColorAttribute({
    raw: themeColor,
    element: border.name ?? "border",
    context,
  });
  if (parsedThemeColor !== undefined) {
    color.themeColor = parsedThemeColor;
  }
  if (themeTint) {
    color.themeTint = themeTint;
  }
  if (themeShade) {
    color.themeShade = themeShade;
  }

  return color;
};

/**
 * `w:val` is `use="required"` on `CT_Border`. An element without it states no
 * style at all, so the border is dropped rather than invented as `none`:
 * `none` is an authored token that cancels an inherited border, and a
 * malformed element is not evidence the author wanted that.
 */
export function parseBorderSpec(
  border: XmlElement | null,
  context?: ParseContext,
): BorderSpec | undefined {
  if (!border) {
    return undefined;
  }

  const rawStyle = getAttribute(border, "w", "val");
  if (!rawStyle) {
    context?.warn({
      code: PARSE_WARNING_CODES.borderWithoutValue,
      element: border.name ?? "border",
    });
    return undefined;
  }

  if (!isBorderStyle(rawStyle)) {
    context?.warn({
      code: PARSE_WARNING_CODES.borderStyleOutsideEnum,
      element: border.name ?? "border",
      value: rawStyle,
    });
  }
  const spec: BorderSpec = { style: borderStyleFrom(rawStyle) };

  const color = parseBorderColor(border, context);
  if (color) {
    spec.color = color;
  }

  const size = parseNumericAttribute(border, "w", "sz");
  if (size !== undefined) {
    spec.size = size;
  }

  const space = parseNumericAttribute(border, "w", "space");
  if (space !== undefined) {
    spec.space = space;
  }

  const shadow = parseOnOffAttribute(border, "w", "shadow");
  if (shadow !== undefined) {
    spec.shadow = shadow;
  }

  const frame = parseOnOffAttribute(border, "w", "frame");
  if (frame !== undefined) {
    spec.frame = frame;
  }

  // Custom page-border art relationship ids. Only `w:pgBorders` carries them,
  // but reading them everywhere costs nothing and keeps one reader: Word
  // re-paints the art glyphs and corner images from these on reload, even
  // though folio renders the underlying line style.
  const artRelationshipId = getAttribute(border, "r", "id")?.trim();
  if (artRelationshipId) {
    spec.artRelationshipId = artRelationshipId;
  }
  const topLeftArtRelationshipId = getAttribute(border, "r", "topLeft")?.trim();
  if (topLeftArtRelationshipId) {
    spec.topLeftArtRelationshipId = topLeftArtRelationshipId;
  }
  const topRightArtRelationshipId = getAttribute(border, "r", "topRight")?.trim();
  if (topRightArtRelationshipId) {
    spec.topRightArtRelationshipId = topRightArtRelationshipId;
  }
  const bottomLeftArtRelationshipId = getAttribute(border, "r", "bottomLeft")?.trim();
  if (bottomLeftArtRelationshipId) {
    spec.bottomLeftArtRelationshipId = bottomLeftArtRelationshipId;
  }
  const bottomRightArtRelationshipId = getAttribute(border, "r", "bottomRight")?.trim();
  if (bottomRightArtRelationshipId) {
    spec.bottomRightArtRelationshipId = bottomRightArtRelationshipId;
  }

  const preservedAttributes = readAttributeBag(border, BORDER_ATTRIBUTES);
  if (preservedAttributes) {
    spec.preservedAttributes = preservedAttributes;
  }

  return spec;
}
