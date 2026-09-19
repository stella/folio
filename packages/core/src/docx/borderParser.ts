/**
 * The one reader for `CT_Border`, shared by the paragraph (`w:pBdr`), style,
 * table (`w:tblBorders`/`w:tcBorders`) and page (`w:pgBorders`) tiers.
 *
 * `w:val` is `ST_Border`, whose 193 members include two distinct "no border"
 * tokens: `nil` and `none`. They are not interchangeable downstream, and an
 * explicit one overrides a border inherited from the container, so the member
 * the author wrote is preserved exactly. Members outside the model's known
 * union survive verbatim rather than collapsing to a default, which is how the
 * repo already treats `w:numFmt`, `w:suff` and `w:tab`.
 */

import type { BorderSpec, ColorValue } from "../types/document";
import { BorderStyleSchema, narrowEnum, ThemeColorSlotSchema } from "./parserEnums";
import { getAttribute, parseNumericAttribute, parseOnOffAttribute } from "./xmlParser";
import type { XmlElement } from "./xmlParser";

const parseBorderColor = (border: XmlElement): ColorValue | undefined => {
  const rgb = getAttribute(border, "w", "color");
  const themeColor = getAttribute(border, "w", "themeColor");
  const themeTint = getAttribute(border, "w", "themeTint");
  const themeShade = getAttribute(border, "w", "themeShade");
  if (rgb === null && themeColor === null && themeTint === null && themeShade === null) {
    return undefined;
  }

  const color: ColorValue = {};
  if (rgb === "auto") {
    color.auto = true;
  } else if (rgb) {
    color.rgb = rgb;
  }

  const validatedThemeColor = narrowEnum(themeColor, ThemeColorSlotSchema);
  if (validatedThemeColor) {
    color.themeColor = validatedThemeColor;
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
export function parseBorderSpec(border: XmlElement | null): BorderSpec | undefined {
  if (!border) {
    return undefined;
  }

  const rawStyle = getAttribute(border, "w", "val");
  if (!rawStyle) {
    return undefined;
  }

  const spec: BorderSpec = { style: narrowEnum(rawStyle, BorderStyleSchema) ?? rawStyle };

  const color = parseBorderColor(border);
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
  const artRelationshipId = getAttribute(border, "w", "id")?.trim();
  if (artRelationshipId) {
    spec.artRelationshipId = artRelationshipId;
  }
  const topLeftArtRelationshipId = getAttribute(border, "w", "topLeft")?.trim();
  if (topLeftArtRelationshipId) {
    spec.topLeftArtRelationshipId = topLeftArtRelationshipId;
  }
  const topRightArtRelationshipId = getAttribute(border, "w", "topRight")?.trim();
  if (topRightArtRelationshipId) {
    spec.topRightArtRelationshipId = topRightArtRelationshipId;
  }
  const bottomLeftArtRelationshipId = getAttribute(border, "w", "bottomLeft")?.trim();
  if (bottomLeftArtRelationshipId) {
    spec.bottomLeftArtRelationshipId = bottomLeftArtRelationshipId;
  }
  const bottomRightArtRelationshipId = getAttribute(border, "w", "bottomRight")?.trim();
  if (bottomRightArtRelationshipId) {
    spec.bottomRightArtRelationshipId = bottomRightArtRelationshipId;
  }

  return spec;
}
