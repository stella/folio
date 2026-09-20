/**
 * Run Serializer - Serialize runs to OOXML XML
 *
 * Converts Run objects back to <w:r> XML format for DOCX files.
 * Handles all formatting properties and content types.
 *
 * OOXML Reference:
 * - Run: w:r
 * - Run properties: w:rPr
 * - Text content: w:t
 */

import { panic } from "better-result";
import type {
  Run,
  RunContent,
  TextContent,
  TabContent,
  BreakContent,
  SymbolContent,
  NoteReferenceContent,
  FieldCharContent,
  InstrTextContent,
  SoftHyphenContent,
  NoBreakHyphenContent,
  DrawingContent,
  ShapeContent,
  TextFormatting,
  ColorValue,
  Image,
  ImageTransform,
  ShapeFill,
  ShapeOutline,
  ShapeTextBody,
  ImagePosition,
  ImageWrap,
  BlockContent,
  RunPropertyChange,
} from "../../types/document";
import { escapeXmlAttribute, escapeXmlText, requiresXmlSpacePreserve } from "@stll/docx-core";
import { isValidHexColor } from "../../utils/colorResolver";
import { normalizeImageLuminancePercent } from "../../utils/imageLuminance";
import { serializePreservedAttributes } from "../attributeRemainder";
import { THEME_COLOR_TO_DRAWING_SCHEME } from "../drawingUtils";
import { fieldStateAttributes } from "../fieldState";
import { serializeGraphicFrameLocks } from "../graphicFrameLocks";
import { canReplayEditableImageRawXml } from "../imageRawXml";
import { serializeNonVisualDrawingNames } from "../nonVisualDrawingProps";
import { DECORATIVE_EXTENSION_URI, DECORATIVE_NAMESPACE } from "../imageParser";
// oxlint-disable-next-line import/no-cycle -- OOXML model is mutually recursive: shape textboxes hold paragraphs, paragraphs hold runs
import { serializeParagraph } from "./paragraphSerializer";
import { serializeTable } from "./tableSerializer";
import { serializeTextFormatting } from "./textFormattingSerializer";
import {
  getSingularRunPropertyChange,
  serializeTrackedChangeAttributes,
} from "./trackedChangeAttributes";
import { intAttr } from "./xmlUtils";

// ============================================================================
// CONSTANTS
// ============================================================================

/**
 * Auto-incrementing counter for generating unique image/shape IDs.
 * Used as a fallback when `image.id` or `shape.id` is undefined (e.g., pasted images).
 * Starts high (100000) to avoid collisions with IDs parsed from existing DOCX content.
 */
let nextAutoId = 100_000;

/**
 * Reset the auto-incrementing ID counter. Call before each serialization pass
 * to keep IDs deterministic across saves.
 */
export function resetAutoIdCounter(): void {
  nextAutoId = 100_000;
}

/** Get a unique positive integer ID, using the provided value or generating one */
function getUniqueId(id: string | number | undefined): string {
  if (id !== undefined && id !== "" && id !== 0) {
    return String(id);
  }
  return String(nextAutoId++);
}

function extractRPrInner(rPrXml: string): string {
  if (!rPrXml.startsWith("<w:rPr>") || !rPrXml.endsWith("</w:rPr>")) {
    return "";
  }
  return rPrXml.slice("<w:rPr>".length, -"</w:rPr>".length);
}

function serializeRunPropertyChange(change: RunPropertyChange): string {
  const previousRPrXml = serializeTextFormatting(change.previousFormatting) || "<w:rPr/>";
  return `<w:rPrChange ${serializeTrackedChangeAttributes(change.info)}>${previousRPrXml}</w:rPrChange>`;
}

function serializeRunProperties(
  formatting: TextFormatting | undefined,
  propertyChanges: RunPropertyChange[] | undefined,
): string {
  const currentRPrXml = serializeTextFormatting(formatting);
  const currentInner = currentRPrXml ? extractRPrInner(currentRPrXml) : "";
  const propertyChange = getSingularRunPropertyChange(propertyChanges);
  const propertyChangeXml = propertyChange ? serializeRunPropertyChange(propertyChange) : "";
  const combined = `${currentInner}${propertyChangeXml}`;

  if (!combined) {
    return "";
  }

  return `<w:rPr>${combined}</w:rPr>`;
}

// ============================================================================
// RUN CONTENT SERIALIZATION
// ============================================================================

/**
 * Serialize text content (w:t)
 */
function serializeTextContent(content: TextContent): string {
  const needsPreserve = requiresXmlSpacePreserve(content.text);

  const spaceAttr = needsPreserve ? ' xml:space="preserve"' : "";

  return `<w:t${spaceAttr}>${escapeXmlText(content.text)}</w:t>`;
}

/**
 * Serialize tab content (w:tab)
 */
function serializeTabContent(content: TabContent): string {
  if (!content.positional) {
    return "<w:tab/>";
  }
  const attrs = [
    content.positional.relativeTo
      ? ` w:relativeTo="${escapeXmlAttribute(content.positional.relativeTo)}"`
      : "",
    content.positional.alignment
      ? ` w:alignment="${escapeXmlAttribute(content.positional.alignment)}"`
      : "",
    content.positional.leader ? ` w:leader="${escapeXmlAttribute(content.positional.leader)}"` : "",
  ].join("");
  return `<w:ptab${attrs}/>`;
}

/**
 * Serialize break content (w:br)
 */
function serializeBreakContent(content: BreakContent): string {
  const attrs: string[] = [];

  if (content.breakType === "page") {
    attrs.push('w:type="page"');
  } else if (content.breakType === "column") {
    attrs.push('w:type="column"');
  } else if (content.breakType === "textWrapping") {
    attrs.push('w:type="textWrapping"');
  }

  if (content.clear !== undefined) {
    attrs.push(`w:clear="${content.clear}"`);
  }

  if (attrs.length === 0) {
    return "<w:br/>";
  }

  return `<w:br ${attrs.join(" ")}/>`;
}

/**
 * Serialize symbol content (w:sym)
 */
function serializeSymbolContent(content: SymbolContent): string {
  // Both attributes are optional on `CT_Sym`. An empty one is the parser's
  // record of an attribute the source did not write, so writing it back as
  // `w:font=""` would invent a value the document never had.
  const font = content.font === "" ? "" : ` w:font="${escapeXmlAttribute(content.font)}"`;
  const char = content.char === "" ? "" : ` w:char="${escapeXmlAttribute(content.char)}"`;
  return `<w:sym${font}${char}/>`;
}

/**
 * Serialize footnote/endnote reference
 */
function serializeNoteReference(content: NoteReferenceContent): string {
  if (content.type === "footnoteRef") {
    return `<w:footnoteReference w:id="${content.id}"/>`;
  }
  return `<w:endnoteReference w:id="${content.id}"/>`;
}

/**
 * Serialize field character (w:fldChar)
 */
function serializeFieldChar(content: FieldCharContent): string {
  const attrs: string[] = [`w:fldCharType="${content.charType}"`, ...fieldStateAttributes(content)];

  return `<w:fldChar ${attrs.join(" ")}/>`;
}

/**
 * Serialize field instruction text (w:instrText)
 */
function serializeInstrText(content: InstrTextContent): string {
  const needsPreserve =
    content.text.startsWith(" ") || content.text.endsWith(" ") || content.text.includes("  ");

  const spaceAttr = needsPreserve ? ' xml:space="preserve"' : "";

  return `<w:instrText${spaceAttr}>${escapeXmlText(content.text)}</w:instrText>`;
}

/**
 * Serialize soft hyphen (w:softHyphen)
 */
function serializeSoftHyphen(_content: SoftHyphenContent): string {
  return "<w:softHyphen/>";
}

/**
 * Serialize non-breaking hyphen (w:noBreakHyphen)
 */
function serializeNoBreakHyphen(_content: NoBreakHyphenContent): string {
  return "<w:noBreakHyphen/>";
}

// ============================================================================
// DRAWING / IMAGE / SHAPE SERIALIZATION
// ============================================================================

/** Serialize a color value to DrawingML a:srgbClr or a:schemeClr */
function serializeDrawingColor(color: ColorValue | undefined): string {
  if (!color) {
    return "";
  }
  if (color.rgb && isValidHexColor(color.rgb)) {
    return `<a:srgbClr val="${escapeXmlAttribute(color.rgb.replace("#", ""))}"/>`;
  }
  if (color.themeColor) {
    const schemeColor = THEME_COLOR_TO_DRAWING_SCHEME[color.themeColor];
    let clr = `<a:schemeClr val="${schemeColor}"`;
    if (color.themeTint) {
      clr += `><a:tint val="${escapeXmlAttribute(color.themeTint)}"/></a:schemeClr>`;
    } else if (color.themeShade) {
      clr += `><a:shade val="${escapeXmlAttribute(color.themeShade)}"/></a:schemeClr>`;
    } else {
      clr += `/>`;
    }
    return clr;
  }
  return "";
}

/** Serialize shape fill to DrawingML */
function serializeFill(fill: ShapeFill | undefined): string {
  if (!fill) {
    return "";
  }
  if (fill.rawXml) {
    return fill.rawXml;
  }
  if (fill.type === "none") {
    return "<a:noFill/>";
  }
  if (fill.type === "solid" && fill.color) {
    return `<a:solidFill>${serializeDrawingColor(fill.color)}</a:solidFill>`;
  }
  if (fill.type === "gradient" && fill.gradient) {
    const g = fill.gradient;
    const stops = g.stops
      .map((s) => `<a:gs pos="${s.position}">${serializeDrawingColor(s.color)}</a:gs>`)
      .join("");
    const direction = (() => {
      if (g.type === "linear") {
        return `<a:lin ang="${(g.angle ?? 0) * 60_000}" scaled="1"/>`;
      }
      let path = "shape";
      if (g.type === "radial") {
        path = "circle";
      } else if (g.type === "rectangular") {
        path = "rect";
      }
      return `<a:path path="${path}"/>`;
    })();
    return `<a:gradFill><a:gsLst>${stops}</a:gsLst>${direction}</a:gradFill>`;
  }
  return "";
}

/** Serialize shape outline to DrawingML a:ln */
function serializeLineCap(cap: NonNullable<ShapeOutline["cap"]>): string {
  if (cap === "round") {
    return "rnd";
  }
  if (cap === "square") {
    return "sq";
  }
  return "flat";
}

function serializeOutline(outline: ShapeOutline | undefined): string {
  if (!outline) {
    return "";
  }
  if (outline.rawXml) {
    return outline.rawXml;
  }
  const attrs: string[] = [];
  if (typeof outline.width === "number") {
    attrs.push(`w="${outline.width}"`);
  }
  if (outline.cap) {
    attrs.push(`cap="${serializeLineCap(outline.cap)}"`);
  }

  const parts: string[] = [];
  if (outline.color) {
    parts.push(`<a:solidFill>${serializeDrawingColor(outline.color)}</a:solidFill>`);
  }
  if (outline.style) {
    parts.push(`<a:prstDash val="${outline.style}"/>`);
  }
  if (outline.join === "bevel") {
    parts.push("<a:bevel/>");
  } else if (outline.join === "round") {
    parts.push("<a:round/>");
  } else if (outline.join === "miter") {
    parts.push("<a:miter/>");
  }
  if (outline.headEnd) {
    parts.push(
      `<a:headEnd type="${outline.headEnd.type}"${outline.headEnd.width ? ` w="${outline.headEnd.width}"` : ""}${outline.headEnd.length ? ` len="${outline.headEnd.length}"` : ""}/>`,
    );
  }
  if (outline.tailEnd) {
    parts.push(
      `<a:tailEnd type="${outline.tailEnd.type}"${outline.tailEnd.width ? ` w="${outline.tailEnd.width}"` : ""}${outline.tailEnd.length ? ` len="${outline.tailEnd.length}"` : ""}/>`,
    );
  }

  if (parts.length === 0 && attrs.length === 0) {
    return "";
  }
  return `<a:ln${attrs.length ? ` ${attrs.join(" ")}` : ""}>${parts.join("")}</a:ln>`;
}

/** Build wp:positionH and wp:positionV for floating drawings */
function serializePosition(pos: ImagePosition): string {
  const parts: string[] = [];

  // Horizontal
  const h = pos.horizontal;
  parts.push(`<wp:positionH relativeFrom="${h.relativeTo}">`);
  if (h.alignment) {
    parts.push(`<wp:align>${h.alignment}</wp:align>`);
  } else {
    parts.push(`<wp:posOffset>${intAttr(h.posOffset)}</wp:posOffset>`);
  }
  parts.push("</wp:positionH>");

  // Vertical
  const v = pos.vertical;
  parts.push(`<wp:positionV relativeFrom="${v.relativeTo}">`);
  if (v.alignment) {
    parts.push(`<wp:align>${v.alignment}</wp:align>`);
  } else {
    parts.push(`<wp:posOffset>${intAttr(v.posOffset)}</wp:posOffset>`);
  }
  parts.push("</wp:positionV>");

  return parts.join("");
}

/** Serialize wrap type to wp:wrap* element */
function serializeWrap(wrap: ImageWrap): string {
  const wrapText = wrap.wrapText ? ` wrapText="${wrap.wrapText}"` : ' wrapText="bothSides"';
  switch (wrap.type) {
    case "square":
      return `<wp:wrapSquare${wrapText}/>`;
    case "tight":
      return `<wp:wrapTight${wrapText}><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>`;
    case "through":
      return `<wp:wrapThrough${wrapText}><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapThrough>`;
    case "topAndBottom":
      return "<wp:wrapTopAndBottom/>";
    case "behind":
    case "inFront":
      return "<wp:wrapNone/>";
    case "inline":
      // Inline images don't get a wrap element — they're laid out
      // inline with text via wp:inline. Caller should not invoke
      // serializeWrap for inline images, but if it slips through we
      // fall back to wrapNone (matches Word's degenerate handling).
      return "<wp:wrapNone/>";
  }
}

/**
 * `a:xfrm` attributes for an authored transform.
 *
 * Written iff authored: `rot="0"` and `flipH="0"` are OOXML's defaults, so a
 * truthiness guard cannot tell "the author said none" from "the author said
 * nothing", and a save erased an authored zero.
 */
const serializeTransformAttrs = (transform: ImageTransform | undefined): string => {
  if (transform === undefined) {
    return "";
  }
  const rot =
    transform.rotation === undefined ? "" : ` rot="${Math.round(transform.rotation * 60_000)}"`;
  const flipH = transform.flipH === undefined ? "" : ` flipH="${transform.flipH ? 1 : 0}"`;
  const flipV = transform.flipV === undefined ? "" : ` flipV="${transform.flipV ? 1 : 0}"`;
  return `${rot}${flipH}${flipV}`;
};

/**
 * Build the common a:graphic > pic:pic element for images.
 *
 * Takes the relationship id as a separate argument so the caller has to have
 * one. There is no default: a picture written against an invented id binds to
 * whichever relationship the part happens to hold under that name.
 */
function serializePicGraphic(image: Image, imageRId: string, sharedId: string): string {
  const cx = image.size.width;
  const cy = image.size.height;
  const rId = escapeXmlAttribute(imageRId);
  const id = sharedId;
  const name = image.filename || `image${id}`;

  const xfrmAttrs = serializeTransformAttrs(image.transform);

  // eigenpal #424: emit <a:srcRect/> for wp:srcRect crop. Each side is a
  // fraction in [0, 1]; OOXML expects 1/100000 units. Zero sides are
  // omitted so the element stays terse for the common case.
  const cropAttrs: string[] = [];
  if (image.crop?.left) {
    cropAttrs.push(`l="${Math.round(image.crop.left * 100_000)}"`);
  }
  if (image.crop?.top) {
    cropAttrs.push(`t="${Math.round(image.crop.top * 100_000)}"`);
  }
  if (image.crop?.right) {
    cropAttrs.push(`r="${Math.round(image.crop.right * 100_000)}"`);
  }
  if (image.crop?.bottom) {
    cropAttrs.push(`b="${Math.round(image.crop.bottom * 100_000)}"`);
  }
  const srcRectEl = cropAttrs.length > 0 ? `<a:srcRect ${cropAttrs.join(" ")}/>` : "";

  const luminanceAttrs: string[] = [];
  if (image.effects?.brightness !== undefined && Number.isFinite(image.effects.brightness)) {
    luminanceAttrs.push(
      `bright="${Math.round(normalizeImageLuminancePercent(image.effects.brightness) * 1_000)}"`,
    );
  }
  if (image.effects?.contrast !== undefined && Number.isFinite(image.effects.contrast)) {
    luminanceAttrs.push(
      `contrast="${Math.round(normalizeImageLuminancePercent(image.effects.contrast) * 1_000)}"`,
    );
  }
  const luminanceChild = luminanceAttrs.length === 0 ? "" : `<a:lum ${luminanceAttrs.join(" ")}/>`;

  // <a:blip> with optional luminance and transparency children.
  // OOXML stores the alpha amount in 1/100000 units. Mirrors eigenpal #424.
  // `image.opacity < 1` is guaranteed by the branch, so only clamp the
  // lower bound.
  const alphaChild =
    image.opacity !== undefined && image.opacity < 1
      ? `<a:alphaModFix amt="${Math.round(Math.max(0, image.opacity) * 100_000)}"/>`
      : "";
  const blipChildren = `${luminanceChild}${alphaChild}`;
  const blipEl = blipChildren
    ? `<a:blip r:embed="${rId}">${blipChildren}</a:blip>`
    : `<a:blip r:embed="${rId}"/>`;

  return [
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">',
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">',
    "<pic:nvPicPr>",
    `<pic:cNvPr id="${id}" name="${escapeXmlAttribute(name)}"${image.alt ? ` descr="${escapeXmlAttribute(image.alt)}"` : ""}/>`,
    "<pic:cNvPicPr/>",
    "</pic:nvPicPr>",
    "<pic:blipFill>",
    blipEl,
    srcRectEl,
    "<a:stretch><a:fillRect/></a:stretch>",
    "</pic:blipFill>",
    "<pic:spPr>",
    `<a:xfrm${xfrmAttrs}>`,
    '<a:off x="0" y="0"/>',
    `<a:ext cx="${intAttr(cx)}" cy="${intAttr(cy)}"/>`,
    "</a:xfrm>",
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
    image.outline ? serializeOutline(image.outline) : "",
    "</pic:spPr>",
    "</pic:pic>",
    "</a:graphicData>",
    "</a:graphic>",
  ].join("");
}

/**
 * Serialize only authored wrap-distance attributes. OOXML defaults omitted
 * values to zero, but preserving absence keeps untouched models stable.
 */
function serializeWrapDistanceAttrs(wrap: ImageWrap | undefined): string {
  const attrs: string[] = [];
  if (wrap?.distT !== undefined) {
    attrs.push(`distT="${intAttr(wrap.distT)}"`);
  }
  if (wrap?.distB !== undefined) {
    attrs.push(`distB="${intAttr(wrap.distB)}"`);
  }
  if (wrap?.distL !== undefined) {
    attrs.push(`distL="${intAttr(wrap.distL)}"`);
  }
  if (wrap?.distR !== undefined) {
    attrs.push(`distR="${intAttr(wrap.distR)}"`);
  }
  return attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
}

/**
 * Serialize the `wp:docPr` extension list.
 *
 * The decorative flag is the one extension folio models; every other `a:ext`
 * the source carried is replayed as it was captured, in its original order,
 * because an extension dropped on save is a fact the document had and lost.
 * The decorative extension goes first: it is the one folio may have added.
 */
function serializeDocPrExtensions(image: Image): string {
  const preserved = image.docPrExtensions ?? [];
  const decorative =
    image.decorative === undefined
      ? ""
      : `<a:ext uri="${DECORATIVE_EXTENSION_URI}"><adec:decorative xmlns:adec="${DECORATIVE_NAMESPACE}" val="${image.decorative ? "1" : "0"}"/></a:ext>`;
  const entries = `${decorative}${preserved.join("")}`;
  return entries ? `<a:extLst>${entries}</a:extLst>` : "";
}

/**
 * Serialize drawing/image content (w:drawing) to full DrawingML XML
 */
function serializeDrawingContent(content: DrawingContent): string {
  const image = content.image;
  const isFloating = image.wrap.type !== "inline";
  const cx = image.size.width;
  const cy = image.size.height;
  // ECMA-376 §20.4.2.8: dist* on wp:inline / wp:anchor are text-wrap
  // distances. Per §20.4.2.5, the visual-effect reservation lives on the
  // separate <wp:effectExtent> element. Don't fold `image.padding`
  // (effectExtent) into wrap dist* — that's the eigenpal #424 fix.
  const wrapDistanceAttrs = serializeWrapDistanceAttrs(image.wrap);
  const effL = image.padding?.left ?? 0;
  const effT = image.padding?.top ?? 0;
  const effR = image.padding?.right ?? 0;
  const effB = image.padding?.bottom ?? 0;
  const effectExtentEl = `<wp:effectExtent l="${intAttr(effL)}" t="${intAttr(effT)}" r="${intAttr(effR)}" b="${intAttr(effB)}"/>`;
  const docPrId = getUniqueId(image.id);
  const docPrNames = serializeNonVisualDrawingNames({
    ...(image.docPrName !== undefined ? { name: image.docPrName } : {}),
    ...(image.alt !== undefined ? { alt: image.alt } : {}),
    ...(image.title !== undefined ? { title: image.title } : {}),
  });
  const hlinkClick = image.hlinkRId
    ? `<a:hlinkClick r:id="${escapeXmlAttribute(image.hlinkRId)}"/>`
    : "";
  // `@hidden` is the drawing not being displayed; it says nothing about
  // whether the image carries information. One `wp:docPr` for both anchorings:
  // an attribute written on one and omitted on the other loses the fact the
  // moment an inline image is anchored, or the reverse.
  const docPrHidden = image.hidden === undefined ? "" : ` hidden="${image.hidden ? "1" : "0"}"`;
  const docPrAttrs = `id="${docPrId}"${docPrNames}${docPrHidden}`;
  const docPrChildren = `${hlinkClick}${serializeDocPrExtensions(image)}`;
  const docPr = docPrChildren
    ? `<wp:docPr ${docPrAttrs}>${docPrChildren}</wp:docPr>`
    : `<wp:docPr ${docPrAttrs}/>`;

  const graphicFramePr = serializeGraphicFrameLocks(image.frameLocks);

  // A drawing with no picture relationship had no `a:blip` to read one from,
  // so it had no `pic:pic` either: a chart, an OLE frame, or an anchor with no
  // graphic at all. The anchor is written back as what it was, without a
  // graphic, rather than upgraded into a picture bound to a borrowed id.
  const graphic = image.rId === undefined ? "" : serializePicGraphic(image, image.rId, docPrId);

  if (!isFloating) {
    // Inline image
    return [
      "<w:drawing>",
      `<wp:inline${wrapDistanceAttrs}>`,
      `<wp:extent cx="${intAttr(cx)}" cy="${intAttr(cy)}"/>`,
      effectExtentEl,
      docPr,
      graphicFramePr,
      graphic,
      "</wp:inline>",
      "</w:drawing>",
    ].join("");
  }

  // Floating (anchored) image
  const behindDoc = image.wrap.type === "behind" ? "1" : "0";
  const position = image.position
    ? serializePosition(image.position)
    : '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>';
  const wrap = serializeWrap(image.wrap);
  // Tri-state: explicit `false` → "0"; explicit `true` or absent →
  // "1" (the OOXML default). Mirrors eigenpal #424.
  const layoutInCellAttr = image.layoutInCell === false ? "0" : "1";
  const allowOverlapAttr = image.allowOverlap === false ? "0" : "1";

  return [
    "<w:drawing>",
    `<wp:anchor${wrapDistanceAttrs} simplePos="0" relativeHeight="251658240" behindDoc="${behindDoc}" locked="0" layoutInCell="${layoutInCellAttr}" allowOverlap="${allowOverlapAttr}">`,
    '<wp:simplePos x="0" y="0"/>',
    position,
    `<wp:extent cx="${intAttr(cx)}" cy="${intAttr(cy)}"/>`,
    effectExtentEl,
    wrap,
    docPr,
    graphicFramePr,
    graphic,
    "</wp:anchor>",
    "</w:drawing>",
  ].join("");
}

/** Serialize text body content for shapes/textboxes */
function serializeShapeTextBody(
  blocks: Extract<BlockContent, { type: "paragraph" | "table" }>[],
): string {
  return blocks
    .map((block) =>
      block.type === "paragraph"
        ? serializeParagraph(block)
        : serializeTable(block, serializeParagraph),
    )
    .join("");
}

const serializeTextBodyAnchor = (anchor: NonNullable<ShapeTextBody["anchor"]>): string => {
  switch (anchor) {
    case "top":
      return "t";
    case "middle":
      return "ctr";
    case "bottom":
      return "b";
    case "distributed":
      return "dist";
    case "justified":
      return "just";
    default:
      return anchor satisfies never;
  }
};

function serializeGeometryAdjustments(shape: ShapeContent["shape"]): string {
  if (!shape.geometryAdjustments || shape.geometryAdjustments.length === 0) {
    return "<a:avLst/>";
  }
  const adjustments = shape.geometryAdjustments
    .map(
      ({ name, formula }) =>
        `<a:gd name="${escapeXmlAttribute(name)}" fmla="${escapeXmlAttribute(formula)}"/>`,
    )
    .join("");
  return `<a:avLst>${adjustments}</a:avLst>`;
}

/**
 * Serialize shape content to full DrawingML XML (wps:wsp inside w:drawing)
 */
function serializeShapeContent(content: ShapeContent): string {
  const shape = content.shape;
  const cx = shape.size.width;
  const cy = shape.size.height;
  const isTextBox = shape.shapeType === "textBox";
  const isFloating = shape.wrap && shape.wrap.type !== "inline";
  const wrapDistances = serializeWrapDistanceAttrs(shape.wrap);
  const docPrId = getUniqueId(shape.id);
  const docPrNames = serializeNonVisualDrawingNames({
    ...(shape.name !== undefined ? { name: shape.name } : {}),
    ...(shape.alt !== undefined ? { alt: shape.alt } : {}),
    ...(shape.title !== undefined ? { title: shape.title } : {}),
  });

  const xfrmAttrs = serializeTransformAttrs(shape.transform);

  // Build wps:spPr
  const spPr = [
    "<wps:spPr>",
    `<a:xfrm${xfrmAttrs}>`,
    '<a:off x="0" y="0"/>',
    `<a:ext cx="${intAttr(cx)}" cy="${intAttr(cy)}"/>`,
    "</a:xfrm>",
    `<a:prstGeom prst="${shape.shapeType === "textBox" ? "rect" : shape.shapeType}">${serializeGeometryAdjustments(shape)}</a:prstGeom>`,
    serializeFill(shape.fill),
    serializeOutline(shape.outline),
    "</wps:spPr>",
  ].join("");

  // Build text body if present
  let textBody = "";
  if (shape.textBody) {
    const tb = shape.textBody;
    const bpAttrs: string[] = ['rot="0"', 'vert="horz"'];
    if (tb.wordArt?.fromWordArt !== undefined) {
      bpAttrs.push(`fromWordArt="${tb.wordArt.fromWordArt ? "1" : "0"}"`);
    }
    if (tb.textWrap) {
      bpAttrs.push(`wrap="${tb.textWrap}"`);
    }
    if (tb.anchor) {
      bpAttrs.push(`anchor="${serializeTextBodyAnchor(tb.anchor)}"`);
    }
    if (tb.anchorCenter) {
      bpAttrs.push('anchorCtr="1"');
    }
    if (tb.margins) {
      if (tb.margins.left != null) {
        bpAttrs.push(`lIns="${intAttr(tb.margins.left)}"`);
      }
      if (tb.margins.top != null) {
        bpAttrs.push(`tIns="${intAttr(tb.margins.top)}"`);
      }
      if (tb.margins.right != null) {
        bpAttrs.push(`rIns="${intAttr(tb.margins.right)}"`);
      }
      if (tb.margins.bottom != null) {
        bpAttrs.push(`bIns="${intAttr(tb.margins.bottom)}"`);
      }
    }

    let autoFitXml = "";
    if (tb.autoFit === "shape") {
      autoFitXml = "<a:spAutoFit/>";
    } else if (tb.autoFit === "normal") {
      autoFitXml = "<a:normAutofit/>";
    } else if (tb.autoFit === "none") {
      autoFitXml = "<a:noAutofit/>";
    }
    const wordArtWarp = tb.wordArt?.preset
      ? `<a:prstTxWarp prst="${escapeXmlAttribute(tb.wordArt.preset)}">${
          tb.wordArt.adjustments && tb.wordArt.adjustments.length > 0
            ? `<a:avLst>${tb.wordArt.adjustments
                .map(
                  ({ name, formula }) =>
                    `<a:gd name="${escapeXmlAttribute(name)}" fmla="${escapeXmlAttribute(formula)}"/>`,
                )
                .join("")}</a:avLst>`
            : "<a:avLst/>"
        }</a:prstTxWarp>`
      : "";
    const bodyPrChildren = `${wordArtWarp}${autoFitXml}`;
    const bodyPrXml = bodyPrChildren
      ? `<wps:bodyPr ${bpAttrs.join(" ")}>${bodyPrChildren}</wps:bodyPr>`
      : `<wps:bodyPr ${bpAttrs.join(" ")}/>`;

    if (isTextBox) {
      textBody = [
        "<wps:txbx><w:txbxContent>",
        serializeShapeTextBody(tb.content),
        "</w:txbxContent></wps:txbx>",
        bodyPrXml,
      ].join("");
    } else {
      textBody = bodyPrXml;
    }
  }

  // `wps:bodyPr` closes the shape's content model whether or not the shape has
  // text, so a shape with no text body still writes an empty one.
  const wsp = [
    "<wps:wsp>",
    `<wps:cNvSpPr${isTextBox ? ' txBox="1"' : ""}/>`,
    spPr,
    textBody === "" ? "<wps:bodyPr/>" : textBody,
    "</wps:wsp>",
  ].join("");

  // Wrap in a:graphic
  const graphic = [
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">',
    wsp,
    "</a:graphicData>",
    "</a:graphic>",
  ].join("");

  if (!isFloating) {
    return [
      "<w:drawing>",
      `<wp:inline${wrapDistances}>`,
      `<wp:extent cx="${intAttr(cx)}" cy="${intAttr(cy)}"/>`,
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
      `<wp:docPr id="${docPrId}"${docPrNames}/>`,
      "<wp:cNvGraphicFramePr/>",
      graphic,
      "</wp:inline>",
      "</w:drawing>",
    ].join("");
  }

  // Floating shape
  const behindDoc = shape.wrap?.type === "behind" ? "1" : "0";
  const position = shape.position
    ? serializePosition(shape.position)
    : '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>';
  if (!shape.wrap) {
    panic("Floating shape must have a wrap property");
  }
  const wrap = serializeWrap(shape.wrap);

  return [
    "<w:drawing>",
    `<wp:anchor${wrapDistances} simplePos="0" relativeHeight="251658240" behindDoc="${behindDoc}" locked="0" layoutInCell="1" allowOverlap="1">`,
    '<wp:simplePos x="0" y="0"/>',
    position,
    `<wp:extent cx="${intAttr(cx)}" cy="${intAttr(cy)}"/>`,
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
    wrap,
    `<wp:docPr id="${docPrId}"${docPrNames}/>`,
    "<wp:cNvGraphicFramePr/>",
    graphic,
    "</wp:anchor>",
    "</w:drawing>",
  ].join("");
}

/**
 * Serialize a single run content item
 */
function serializeRunContent(content: RunContent): string {
  switch (content.type) {
    case "text":
      return serializeTextContent(content);
    case "tab":
      return serializeTabContent(content);
    case "break":
      return serializeBreakContent(content);
    case "symbol":
      return serializeSymbolContent(content);
    case "footnoteRef":
    case "endnoteRef":
      return serializeNoteReference(content);
    case "fieldChar":
      return serializeFieldChar(content);
    case "instrText":
      return serializeInstrText(content);
    case "softHyphen":
      return serializeSoftHyphen(content);
    case "noBreakHyphen":
      return serializeNoBreakHyphen(content);
    case "renderedPageBreak":
      return "<w:lastRenderedPageBreak/>";
    case "preservedXml":
      return content.xml;
    case "drawing":
      if (content.rawXml && canReplayEditableImageRawXml(content)) {
        return content.rawXml;
      }
      return serializeDrawingContent(content);
    case "shape":
      return serializeShapeContent(content);
    default:
      return "";
  }
}

// ============================================================================
// MAIN SERIALIZATION
// ============================================================================

/**
 * Serialize a run to OOXML XML (w:r)
 *
 * @param run - The run to serialize
 * @returns XML string for the run
 */
export function serializeRun(run: Run): string {
  const parts: string[] = [];

  // Add run properties if present
  const rPrXml = serializeRunProperties(run.formatting, run.propertyChanges);
  if (rPrXml) {
    parts.push(rPrXml);
  }

  // Add run content
  for (const content of run.content) {
    const contentXml = serializeRunContent(content);
    if (contentXml) {
      parts.push(contentXml);
    }
  }

  // The run models no attribute of its own, so every one it writes comes from
  // the remainder the parser kept.
  const attrs = serializePreservedAttributes([], run.preservedAttributes);
  return `<w:r${attrs.length > 0 ? ` ${attrs.join(" ")}` : ""}>${parts.join("")}</w:r>`;
}

/**
 * Serialize multiple runs to OOXML XML
 *
 * @param runs - The runs to serialize
 * @returns XML string for all runs
 */
export function serializeRuns(runs: Run[]): string {
  return runs.map(serializeRun).join("");
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Check if a run has any content
 */
export function hasRunContent(run: Run): boolean {
  return run.content.length > 0;
}

/**
 * Check if a run has formatting
 */
export function hasRunFormatting(run: Run): boolean {
  return run.formatting !== undefined && Object.keys(run.formatting).length > 0;
}

/**
 * Get plain text from a run (for comparison/debugging)
 */
export function getRunPlainText(run: Run): string {
  return run.content
    .filter((c): c is TextContent => c.type === "text")
    .map((c) => c.text)
    .join("");
}

/**
 * Create an empty run
 */
export function createEmptyRun(): Run {
  return {
    type: "run",
    content: [],
  };
}

/**
 * Create a text run
 */
export function createTextRun(text: string, formatting?: TextFormatting): Run {
  return {
    type: "run",
    ...(formatting !== undefined ? { formatting } : {}),
    content: [{ type: "text", text }],
  };
}

/**
 * Create a break run
 */
export function createBreakRun(
  breakType?: "page" | "column" | "textWrapping",
  formatting?: TextFormatting,
): Run {
  return {
    type: "run",
    ...(formatting !== undefined ? { formatting } : {}),
    content: [
      {
        type: "break" as const,
        ...(breakType !== undefined ? { breakType } : {}),
      },
    ],
  };
}

/**
 * Create a tab run
 */
export function createTabRun(formatting?: TextFormatting): Run {
  return {
    type: "run",
    ...(formatting !== undefined ? { formatting } : {}),
    content: [{ type: "tab" }],
  };
}
