/**
 * Shared single-border serializer for every `CT_Border` field — table borders
 * (`w:tblBorders`/`w:tcBorders`), paragraph borders (`w:pBdr`), and page borders
 * (`w:pgBorders`). The per-container helpers in the table/paragraph/section
 * serializers all delegate here so a border rule round-trips identically
 * everywhere; previously each carried its own copy and the same bug three times
 * (eigenpal/docx-editor#959).
 */

import type { BorderSpec, ExhaustiveFields } from "../../types/document";
import { intAttr } from "./xmlUtils";
import { escapeXmlAttribute } from "@stll/docx-core";
import { borderStyleToken, themeColorToken } from "@stll/docx-core/model";

type ClassifiedBorderField =
  | "style"
  | "color"
  | "size"
  | "space"
  | "shadow"
  | "frame"
  | "artRelationshipId"
  | "topLeftArtRelationshipId"
  | "topRightArtRelationshipId"
  | "bottomLeftArtRelationshipId"
  | "bottomRightArtRelationshipId";
type ExhaustiveBorder = ExhaustiveFields<BorderSpec, ClassifiedBorderField>;

type BorderColor = NonNullable<BorderSpec["color"]>;
type ClassifiedBorderColorField = "auto" | "rgb" | "themeColor" | "themeTint" | "themeShade";
type ExhaustiveBorderColor = ExhaustiveFields<BorderColor, ClassifiedBorderColorField>;

/**
 * Serialize a single border element (`<w:top .../>`, `<w:left .../>`, ...).
 *
 * `nil`/`none` mean "no border", but an *explicit* one overrides an inherited
 * border (a table-level grid via `w:tblBorders`, a paragraph-style border, a
 * section page border), so it must still round-trip as `<w:side w:val="nil"/>`.
 * A `BorderSpec` only reaches here when the source set it or the user turned the
 * border off, so emitting it is faithful, not noise — dropping it silently
 * re-inherited the container default (e.g. hidden table gridlines reappeared as
 * a full grid on reload). A `nil`/`none` side usually carries no
 * size/color/space, but Word writes `w:sz="0" w:space="0" w:color="auto"` on a
 * turned-off side; those are emitted when present so the value survives a
 * save→parse round-trip and is not otherwise added.
 *
 * An unrecognised `w:val` and the color values come straight from the parsed
 * DOCX, so they are untrusted and are `escapeXmlAttribute`'d before re-entering
 * XML attributes; for valid documents these are enum/hex values, so escaping is
 * a no-op.
 */
export function serializeBorder(input: ExhaustiveBorder | undefined, elementName: string): string {
  if (!input) {
    return "";
  }

  const border: ExhaustiveBorder = input;

  const {
    style,
    color,
    size,
    space,
    shadow,
    frame,
    artRelationshipId,
    topLeftArtRelationshipId,
    topRightArtRelationshipId,
    bottomLeftArtRelationshipId,
    bottomRightArtRelationshipId,
  } = border;

  const attrs: string[] = [`w:val="${escapeXmlAttribute(borderStyleToken(style))}"`];

  if (size !== undefined) {
    attrs.push(`w:sz="${intAttr(size)}"`);
  }

  if (space !== undefined) {
    attrs.push(`w:space="${intAttr(space)}"`);
  }

  if (color) {
    const exhaustiveColor: ExhaustiveBorderColor = color;
    const { auto, rgb, themeColor, themeTint, themeShade } = exhaustiveColor;
    if (auto) {
      attrs.push('w:color="auto"');
    } else if (rgb) {
      attrs.push(`w:color="${escapeXmlAttribute(rgb)}"`);
    }

    if (themeColor) {
      attrs.push(`w:themeColor="${escapeXmlAttribute(themeColorToken(themeColor))}"`);
    }

    if (themeTint) {
      attrs.push(`w:themeTint="${escapeXmlAttribute(themeTint)}"`);
    }

    if (themeShade) {
      attrs.push(`w:themeShade="${escapeXmlAttribute(themeShade)}"`);
    }
  }

  // `ST_OnOff` has three states and `CT_Border` gives neither attribute an XSD
  // default, so an explicit off is not an absence: write whatever the border
  // authored, and write nothing when it authored nothing. `1`/`0` is what Word
  // writes and what the table and section serializers already write.
  if (shadow !== undefined) {
    attrs.push(`w:shadow="${shadow ? "1" : "0"}"`);
  }

  if (frame !== undefined) {
    attrs.push(`w:frame="${frame ? "1" : "0"}"`);
  }

  // Custom page-border art relationship ids (only present on `w:pgBorders`
  // sides; undefined for table/paragraph borders, so skipped there). They are
  // relationship references, so they live in the `r` namespace
  // (`CT_PageBorder`, `CT_TopPageBorder`, `CT_BottomPageBorder`); a `w:id`
  // here is a different attribute, which Word drops and the art with it.
  if (artRelationshipId) {
    attrs.push(`r:id="${escapeXmlAttribute(artRelationshipId)}"`);
  }

  if (topLeftArtRelationshipId) {
    attrs.push(`r:topLeft="${escapeXmlAttribute(topLeftArtRelationshipId)}"`);
  }

  if (topRightArtRelationshipId) {
    attrs.push(`r:topRight="${escapeXmlAttribute(topRightArtRelationshipId)}"`);
  }

  if (bottomLeftArtRelationshipId) {
    attrs.push(`r:bottomLeft="${escapeXmlAttribute(bottomLeftArtRelationshipId)}"`);
  }

  if (bottomRightArtRelationshipId) {
    attrs.push(`r:bottomRight="${escapeXmlAttribute(bottomRightArtRelationshipId)}"`);
  }

  return `<w:${elementName} ${attrs.join(" ")}/>`;
}
