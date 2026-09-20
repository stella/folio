/**
 * Image Parser - Parse embedded images from w:drawing elements
 *
 * DOCX images are contained in <w:drawing> elements with either:
 * - wp:inline - Inline images that flow with text
 * - wp:anchor - Floating/anchored images with text wrapping
 *
 * OOXML Structure:
 * w:drawing
 *   ├── wp:inline or wp:anchor
 *   │   ├── wp:extent (size: cx, cy in EMUs)
 *   │   ├── wp:effectExtent (effect margins)
 *   │   ├── wp:docPr (document properties: id, name, descr, title)
 *   │   ├── wp:positionH / wp:positionV (for anchor only)
 *   │   ├── wp:wrap* (wrapping mode for anchor: wrapNone, wrapSquare, etc.)
 *   │   └── a:graphic
 *   │       └── a:graphicData
 *   │           └── pic:pic
 *   │               ├── pic:nvPicPr (non-visual properties)
 *   │               ├── pic:blipFill
 *   │               │   └── a:blip (r:embed = rId)
 *   │               └── pic:spPr
 *   │                   └── a:xfrm (transform: rotation, flip)
 *
 * EMU (English Metric Units): 914400 EMU = 1 inch
 * Conversion: pixels = (emu * 96) / 914400
 */

import { relationshipIdOf } from "@stll/docx-core/model";

import type {
  Image,
  ImageCrop,
  ImageDocPrLink,
  ImageSize,
  ImagePosition,
  ImageTransform,
  RelationshipId,
  RelationshipMap,
  MediaFile,
} from "../types/document";
import { emuToPixels } from "../utils/units";
import { sanitizeExternalUrl } from "../utils/urlSecurity";
import { sanitizeImageSrc } from "../utils/sanitizeImageSrc";
import { normalizeImageLuminancePercent } from "../utils/imageLuminance";
import {
  parseAnchorBehindDoc,
  parsePositionH,
  parsePositionV,
  findWrapElement,
  parseDrawingEffectExtent,
  parseInlineWrap,
  parseWrapElement,
} from "./drawingUtils";
import { parseDrawingAnchor, WORDPROCESSING_DRAWING_NAMESPACE_URIS } from "./drawingAnchor";
import { parseGraphicFrameLocks } from "./graphicFrameLocks";
import { parseNonVisualDrawingNames } from "./nonVisualDrawingProps";
import { RELATIONSHIP_TYPES, resolveRelationshipIdOfType } from "./relsParser";
import { isTextBoxDrawing } from "./textBoxParser";
import { captureVerbatimXml } from "./verbatimCapture";
import { percentageSpelling } from "./transitionalSpelling";
import {
  findChildByNamespaceUri,
  findChildrenByNamespaceUri,
  getChildElements,
  getAttribute,
  getLocalName,
  getNamespaceUri,
  parseNumericAttribute,
  parseOnOffAttribute,
  parseOnOffValue,
  findByFullName,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

const DRAWINGML_MAIN_NAMESPACE_URIS: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/main",
  "http://purl.oclc.org/ooxml/drawingml/main",
]);
const FIXED_PERCENTAGE = /^-?\d+$/u;

// ============================================================================
// ROTATION CONVERSION
// ============================================================================

/**
 * Convert rotation value (1/60000 of a degree) to degrees
 *
 * @param rot - Rotation in 60000ths of a degree
 * @returns Rotation in degrees
 */
function rotToDegrees(rot: string | null | undefined): number | undefined {
  if (rot === null || rot === undefined) {
    return undefined;
  }
  const val = Number.parseInt(rot, 10);
  if (Number.isNaN(val)) {
    return undefined;
  }
  return val / 60_000;
}

// ============================================================================
// ELEMENT FINDERS
// ============================================================================

/**
 * A child of `wp:inline` or `wp:anchor`, resolved by namespace.
 *
 * The `wp` prefix is the producer's choice, not the document's meaning: a
 * package that binds the WordprocessingDrawing namespace to another prefix
 * carries the same `wp:extent` and a prefix-matched read finds none of it.
 */
const findDrawingChild = (parent: XmlElement | null, localName: string): XmlElement | null =>
  findChildByNamespaceUri(parent, WORDPROCESSING_DRAWING_NAMESPACE_URIS, localName);

// ============================================================================
// SIZE PARSING
// ============================================================================

/**
 * Parse extent element for image size
 *
 * @param extent - wp:extent element
 * @returns ImageSize in EMUs
 */
function parseExtent(extent: XmlElement | null): ImageSize {
  if (!extent) {
    return { width: 0, height: 0 };
  }

  const cx = parseNumericAttribute(extent, null, "cx") ?? 0;
  const cy = parseNumericAttribute(extent, null, "cy") ?? 0;

  return { width: cx, height: cy };
}

// ============================================================================
// DOCUMENT PROPERTIES PARSING
// ============================================================================

/**
 * The `a:ext` uri under which Word records that a drawing is decorative. The
 * extension holds `<adec:decorative val="…"/>`; `CT_NonVisualDrawingProps` has
 * no `@decorative` attribute for it to be written as.
 */
export const DECORATIVE_EXTENSION_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

/** The namespace the decorative extension's element is bound to. */
export const DECORATIVE_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2017/decorative";

/**
 * `wp:docPr`'s extension list is DrawingML's (`a:extLst` holding `a:ext`), in
 * the Transitional or the Strict namespace. Matching on the local name alone
 * would take an `extLst` some other namespace owns and replay its children
 * inside an `a:extLst`, which is a different container than the source wrote.
 */
const DRAWINGML_NAMESPACE_URIS = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/main",
  "http://purl.oclc.org/ooxml/drawingml/main",
]);

type DocPropsExtensions = {
  decorative?: boolean;
  /** Every other `a:ext`, verbatim and in source order. */
  other: string[];
};

/**
 * Read the `wp:docPr` extension list.
 *
 * Only the decorative extension is modeled. The rest — a creation id, a
 * local-DPI hint, whatever a later Word writes — are captured as they were
 * written so the save path can put them back: an extension folio drops is a
 * fact the document had and no longer does.
 */
const parseDocPropsExtensions = (docPr: XmlElement): DocPropsExtensions => {
  const extLst = findChildByNamespaceUri(docPr, DRAWINGML_NAMESPACE_URIS, "extLst");
  if (!extLst) {
    return { other: [] };
  }

  const result: DocPropsExtensions = { other: [] };
  for (const ext of findChildrenByNamespaceUri(extLst, DRAWINGML_NAMESPACE_URIS, "ext")) {
    const uri = getAttribute(ext, null, "uri");
    if (uri?.toUpperCase() !== DECORATIVE_EXTENSION_URI) {
      result.other.push(captureVerbatimXml(ext));
      continue;
    }
    // Resolved namespace, not the prefix: `adec` is the prefix Word writes and
    // a producer is free to pick another for the same URI.
    const flag = getChildElements(ext).find(
      (child) =>
        getLocalName(child.name) === "decorative" &&
        getNamespaceUri(child) === DECORATIVE_NAMESPACE,
    );
    // An extension whose body is not the element it exists for is not an
    // opinion about decorativeness; keep it rather than reading it wrong.
    if (!flag) {
      result.other.push(captureVerbatimXml(ext));
      continue;
    }
    result.decorative = parseOnOffValue(getAttribute(flag, null, "val")) ?? true;
  }
  return result;
};

/**
 * One of `wp:docPr`'s two `CT_Hyperlink` children, as the source wrote it.
 *
 * Resolved by namespace URI rather than by the `a:` prefix: a producer is free
 * to bind DrawingML's namespace to another prefix, and a prefix-matched read
 * would drop the whole link for a document that does.
 */
const parseDocPrLink = (docPr: XmlElement, localName: string): ImageDocPrLink | undefined => {
  const element = findChildByNamespaceUri(docPr, DRAWINGML_NAMESPACE_URIS, localName);
  if (!element) {
    return undefined;
  }
  const rId = getAttribute(element, "r", "id");
  return { xml: captureVerbatimXml(element), ...(rId == null ? {} : { rId }) };
};

/**
 * Parse document properties (wp:docPr)
 *
 * @param docPr - wp:docPr element
 * @returns Object with id, name, description, title
 */
function parseDocProps(docPr: XmlElement | null): {
  id?: string;
  name?: string;
  alt?: string;
  title?: string;
  decorative?: boolean;
  hidden?: boolean;
  docPrExtensions?: string[];
  hlinkClickSource?: ImageDocPrLink;
  hlinkHoverSource?: ImageDocPrLink;
} {
  if (!docPr) {
    return {};
  }

  const id = getAttribute(docPr, null, "id");
  const names = parseNonVisualDrawingNames(docPr);

  // Two separate facts. `@hidden` says the drawing is not displayed;
  // the decorative extension says it is displayed and carries nothing a
  // reader needs. Writing either as the other inverts what the document said.
  const hidden = parseOnOffAttribute(docPr, null, "hidden");
  const { decorative, other: docPrExtensions } = parseDocPropsExtensions(docPr);

  // Check for hyperlink (a:hlinkClick) — clickable image
  const hlinkClickSource = parseDocPrLink(docPr, "hlinkClick");
  const hlinkHoverSource = parseDocPrLink(docPr, "hlinkHover");

  return {
    ...(id != null ? { id } : {}),
    ...names,
    ...(decorative === undefined ? {} : { decorative }),
    ...(hidden === undefined ? {} : { hidden }),
    ...(docPrExtensions.length > 0 ? { docPrExtensions } : {}),
    ...(hlinkClickSource === undefined ? {} : { hlinkClickSource }),
    ...(hlinkHoverSource === undefined ? {} : { hlinkHoverSource }),
  };
}

/**
 * The `wp:docPr` links an image carries back to the save, target-checked.
 *
 * A link's `r:id` has to name a hyperlink relationship with a target
 * `sanitizeExternalUrl` accepts, or folio does not write the element back:
 * replaying the source bytes for a `javascript:` target would reinstate the
 * link the modelled path already refuses. A link with no `r:id` at all — one
 * that carries only a `tooltip`, or a `ppaction://` — names no relationship
 * and so has nothing to check.
 */
type DocPrLinkOptions = {
  props: { hlinkClickSource?: ImageDocPrLink; hlinkHoverSource?: ImageDocPrLink };
  rels: RelationshipMap | undefined;
};

const safeDocPrLinks = ({
  props,
  rels,
}: DocPrLinkOptions): Pick<
  Image,
  "hlinkHref" | "hlinkRId" | "hlinkClickSource" | "hlinkHoverXml"
> => {
  const targetOf = (link: ImageDocPrLink | undefined): string | undefined => {
    if (link === undefined) {
      return undefined;
    }
    if (link.rId === undefined) {
      return "";
    }
    const resolved = resolveRelationshipIdOfType(rels, link.rId, RELATIONSHIP_TYPES.hyperlink);
    return resolved.status === "resolved"
      ? sanitizeExternalUrl(resolved.relationship.target)
      : undefined;
  };

  const click = props.hlinkClickSource;
  const clickHref = targetOf(click);
  const hoverKept = targetOf(props.hlinkHoverSource) !== undefined;
  return {
    ...(clickHref ? { hlinkHref: clickHref } : {}),
    ...(clickHref !== undefined && click?.rId !== undefined ? { hlinkRId: click.rId } : {}),
    ...(clickHref === undefined || click === undefined ? {} : { hlinkClickSource: click }),
    ...(hoverKept && props.hlinkHoverSource !== undefined
      ? { hlinkHoverXml: props.hlinkHoverSource.xml }
      : {}),
  };
};

// ============================================================================
// TRANSFORM PARSING
// ============================================================================

/**
 * Parse transform properties from a:xfrm
 */
function parseTransform(xfrm: XmlElement | null): ImageTransform | undefined {
  if (!xfrm) {
    return undefined;
  }

  // Authored, not truthy: `rot="0"` and `flipH="0"` are the defaults, so a
  // truthiness test cannot tell "the author said none" from "the author said
  // nothing", and the difference has to survive to the save.
  const flipH = parseOnOffValue(getAttribute(xfrm, null, "flipH"));
  const flipV = parseOnOffValue(getAttribute(xfrm, null, "flipV"));

  const rotation = rotToDegrees(getAttribute(xfrm, null, "rot"));

  if (rotation === undefined && flipH === undefined && flipV === undefined) {
    return undefined;
  }

  const transform: ImageTransform = {};
  if (rotation !== undefined) {
    transform.rotation = rotation;
  }
  if (flipH !== undefined) {
    transform.flipH = flipH;
  }
  if (flipV !== undefined) {
    transform.flipV = flipV;
  }

  return transform;
}

// ============================================================================
// BLIP EXTRACTION (image relationship ID)
// ============================================================================

/**
 * Find the pic:blipFill element in a w:drawing container.
 *
 * Path: a:graphic > a:graphicData > pic:pic > pic:blipFill
 *
 * The blipFill carries both `a:blip` (the relationship ID) and the optional
 * `a:srcRect` crop element, so callers that need either share this walk.
 */
function findBlipFillElement(container: XmlElement): XmlElement | null {
  return findByFullName(findPictureElement(container), "pic:blipFill");
}

/** `a:graphic > a:graphicData > pic:pic`, the one walk every picture read shares. */
function findPictureElement(container: XmlElement): XmlElement | null {
  const graphicData = findByFullName(findByFullName(container, "a:graphic"), "a:graphicData");
  return findByFullName(graphicData, "pic:pic");
}

/**
 * `pic:cNvPr` — the picture's own name, alt text and title.
 *
 * Not `wp:docPr`'s. They are two `CT_NonVisualDrawingProps` elements on the
 * same drawing and a reader names the object from whichever one it is looking
 * at, so folding them together loses whichever the source did not repeat.
 */
function findPictureNonVisualProps(container: XmlElement): XmlElement | null {
  const pic = findPictureElement(container);
  return pic ? findByFullName(findByFullName(pic, "pic:nvPicPr"), "pic:cNvPr") : null;
}

/**
 * Parse `<a:srcRect l="..." t="..." r="..." b="..."/>` inside `pic:blipFill`.
 * Values are in 1/100000 of the source image dimension; converted to fractions
 * in [0, 1] so the renderer can apply them as CSS clip-path percentages.
 *
 * eigenpal #424 (image-crop subset).
 */
function parseImageCrop(blipFill: XmlElement | null): ImageCrop | undefined {
  if (!blipFill) {
    return undefined;
  }
  const srcRect = findByFullName(blipFill, "a:srcRect");
  if (!srcRect) {
    return undefined;
  }
  const toFraction = (attr: string): number | undefined => {
    const raw = parseNumericAttribute(srcRect, null, attr);
    if (raw === undefined || raw === 0) {
      return undefined;
    }
    return raw / 100_000;
  };
  const left = toFraction("l");
  const top = toFraction("t");
  const right = toFraction("r");
  const bottom = toFraction("b");
  if (left === undefined && top === undefined && right === undefined && bottom === undefined) {
    return undefined;
  }
  const crop: ImageCrop = {};
  if (left !== undefined) {
    crop.left = left;
  }
  if (top !== undefined) {
    crop.top = top;
  }
  if (right !== undefined) {
    crop.right = right;
  }
  if (bottom !== undefined) {
    crop.bottom = bottom;
  }
  return crop;
}

/**
 * Parse `<a:alphaModFix amt="..."/>` inside the `a:blip` element. The
 * `amt` value is in 1/100000; convert to a fraction in [0, 1] for CSS
 * `opacity`. Returns undefined when no alpha modifier is present (fully
 * opaque), when `amt` is missing or non-numeric, or when `amt` >= 100000
 * (also fully opaque). `parseNumericAttribute` already returns
 * `undefined` (not `NaN`) for non-numeric values, so a downstream NaN
 * is impossible here.
 *
 * Mirrors eigenpal docx-editor #424.
 */
function parseImageOpacity(blip: XmlElement | null): number | undefined {
  if (!blip) {
    return undefined;
  }
  const alpha = findByFullName(blip, "a:alphaModFix");
  if (!alpha) {
    return undefined;
  }
  const amt = parseNumericAttribute(alpha, null, "amt");
  if (amt === undefined || amt >= 100_000) {
    return undefined;
  }
  // `amt < 100_000` is guaranteed above, so the result is < 1; only
  // clamp the lower bound.
  return Math.max(0, amt / 100_000);
}

const parseLuminancePercent = (
  luminance: XmlElement,
  attribute: "bright" | "contrast",
): number | undefined => {
  const raw = getAttribute(luminance, null, attribute);
  if (raw === null) {
    return undefined;
  }
  const strictPercent = percentageSpelling(raw);
  const percent =
    strictPercent ?? (FIXED_PERCENTAGE.test(raw.trim()) ? Number(raw) / 1_000 : undefined);
  return percent === undefined || !Number.isFinite(percent)
    ? undefined
    : normalizeImageLuminancePercent(percent);
};

/** Parse DrawingML `a:lum` into signed percentage values, preserving explicit zeroes. */
function parseImageLuminance(blip: XmlElement | null): Image["effects"] | undefined {
  const luminance = findChildByNamespaceUri(blip, DRAWINGML_MAIN_NAMESPACE_URIS, "lum");
  if (!luminance) {
    return undefined;
  }
  const brightness = parseLuminancePercent(luminance, "bright");
  const contrast = parseLuminancePercent(luminance, "contrast");
  if (brightness === undefined && contrast === undefined) {
    return undefined;
  }
  return {
    ...(brightness === undefined ? {} : { brightness }),
    ...(contrast === undefined ? {} : { contrast }),
  };
}

/**
 * Extract rId from a:blip element.
 *
 * Undefined when the drawing has no blip to read one from: a chart, a diagram
 * or an OLE frame carries an `a:graphic` that is not a picture, and a
 * `wp:inline` may carry no graphic at all.
 */
function extractBlipRId(blip: XmlElement | null): RelationshipId | undefined {
  if (!blip) {
    return undefined;
  }

  // `r:embed` is what Word writes. Some generators drop the prefix, and a
  // linked rather than embedded picture names its target with `r:link`. An
  // attribute present and empty names nothing, so it falls through like one
  // that is not there.
  return (
    relationshipIdOf(getAttribute(blip, "r", "embed")) ??
    relationshipIdOf(getAttribute(blip, null, "embed")) ??
    relationshipIdOf(getAttribute(blip, "r", "link"))
  );
}

/**
 * Find transform (a:xfrm) from picture shape properties
 *
 * Path: a:graphic > a:graphicData > pic:pic > pic:spPr > a:xfrm
 */
function findPictureTransform(container: XmlElement): XmlElement | null {
  const spPr = findByFullName(findPictureElement(container), "pic:spPr");
  return findByFullName(spPr, "a:xfrm");
}

// ============================================================================
// MEDIA RESOLUTION
// ============================================================================

/**
 * Normalize a target path to the standard word/media/... format
 */
function normalizeMediaPath(targetPath: string): string {
  if (!targetPath) {
    return targetPath;
  }

  // Remove leading slashes
  let normalized = targetPath.replace(/^\/+/u, "");

  // Ensure word/ prefix for media files
  if (!normalized.startsWith("word/")) {
    normalized = `word/${normalized}`;
  }

  return normalized;
}

/**
 * Get MIME type from file extension
 */
function getMimeType(path: string): string {
  const ext = path.split(".").pop()?.toLowerCase() ?? "";

  const mimeTypes: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    tiff: "image/tiff",
    tif: "image/tiff",
    webp: "image/webp",
    svg: "image/svg+xml",
    emf: "image/x-emf",
    wmf: "image/x-wmf",
  };

  return mimeTypes[ext] ?? "application/octet-stream";
}

/**
 * Resolve image data from relationships and media map
 *
 * @param rId - Relationship ID (e.g., "rId1")
 * @param rels - Relationship map
 * @param media - Media files map
 * @returns Object with src (data URL or blob), mimeType, and filename
 */
export function resolveImageData(
  rId: string | undefined,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined,
): { src?: string; mimeType?: string; filename?: string } {
  // An id that names no image relationship resolves to nothing. Reading any
  // other part's target as media is how an absent picture comes back named
  // after whichever relationship happened to be first in the part.
  const resolved = resolveRelationshipIdOfType(rels, rId, RELATIONSHIP_TYPES.image);
  if (resolved.status !== "resolved") {
    return {};
  }

  // Get the target path
  const targetPath = resolved.relationship.target;
  if (!targetPath) {
    return {};
  }

  // Normalize the path
  const normalizedPath = normalizeMediaPath(targetPath);
  const filename = targetPath.split("/").pop();

  // Case-insensitive lookup helper for media map
  const findMediaCaseInsensitive = (
    map: Map<string, MediaFile>,
    searchPath: string,
  ): MediaFile | undefined => {
    const lowerPath = searchPath.toLowerCase();
    for (const [key, value] of map.entries()) {
      if (key.toLowerCase() === lowerPath) {
        return value;
      }
    }
    return undefined;
  };

  // Try to find the media file (case-insensitive)
  if (media) {
    // Try normalized path first
    const mediaFile = findMediaCaseInsensitive(media, normalizedPath);
    if (mediaFile) {
      const src = mediaFile.dataUrl || mediaFile.base64;
      return {
        ...(src !== undefined ? { src } : {}),
        mimeType: mediaFile.mimeType,
        ...(filename !== undefined ? { filename } : {}),
      };
    }

    // Try without word/ prefix
    const altPath = targetPath.replace(/^\/+/u, "");
    const altMediaFile = findMediaCaseInsensitive(media, altPath);
    if (altMediaFile) {
      const src = altMediaFile.dataUrl || altMediaFile.base64;
      return {
        ...(src !== undefined ? { src } : {}),
        mimeType: altMediaFile.mimeType,
        ...(filename !== undefined ? { filename } : {}),
      };
    }

    // Try with word/ prefix added
    const withWordPrefix = `word/${altPath}`;
    const prefixedMediaFile = findMediaCaseInsensitive(media, withWordPrefix);
    if (prefixedMediaFile) {
      const src = prefixedMediaFile.dataUrl || prefixedMediaFile.base64;
      return {
        ...(src !== undefined ? { src } : {}),
        mimeType: prefixedMediaFile.mimeType,
        ...(filename !== undefined ? { filename } : {}),
      };
    }
  }

  // Return at least the MIME type based on extension
  return {
    mimeType: getMimeType(targetPath),
    ...(filename !== undefined ? { filename } : {}),
  };
}

// ============================================================================
// MAIN PARSING FUNCTIONS
// ============================================================================

/**
 * Parse a wp:inline element (inline image)
 *
 * @param inlineEl - The wp:inline element
 * @param rels - Relationship map for resolving rId
 * @param media - Media files map
 * @returns Parsed Image object
 */
function parseInline(
  inlineEl: XmlElement,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined,
): Image {
  // Parse extent (size)
  const extent = findDrawingChild(inlineEl, "extent");
  const size = parseExtent(extent);

  // Parse effect extent
  const effectExtent = findDrawingChild(inlineEl, "effectExtent");
  const padding = parseDrawingEffectExtent(effectExtent);

  // Parse document properties
  const docPr = findDrawingChild(inlineEl, "docPr");
  const props = parseDocProps(docPr);

  const frameLocks = parseGraphicFrameLocks(inlineEl);

  // `pic:cNvPr`'s own set, kept apart from `wp:docPr`'s above: the rebuild
  // wrote the media filename here, renaming every picture the author had named.
  const pictureNames = parseNonVisualDrawingNames(findPictureNonVisualProps(inlineEl));

  // Find blip and extract rId
  const blipFill = findBlipFillElement(inlineEl);
  const blip = blipFill ? findByFullName(blipFill, "a:blip") : null;
  const rId = extractBlipRId(blip);
  const crop = parseImageCrop(blipFill);
  const opacity = parseImageOpacity(blip);
  const effects = parseImageLuminance(blip);

  // Resolve image data
  const imageData = resolveImageData(rId, rels, media);

  // Find transform
  const xfrm = findPictureTransform(inlineEl);
  const transform = parseTransform(xfrm);

  const wrap = parseInlineWrap(inlineEl);

  const image: Image = {
    type: "image",
    ...(rId === undefined ? {} : { rId }),
    size,
    wrap,
  };

  // Add optional properties
  if (props.id) {
    image.id = props.id;
  }
  if (props.name !== undefined) {
    image.docPrName = props.name;
  }
  if (props.alt !== undefined) {
    image.alt = props.alt;
  }
  if (props.title !== undefined) {
    image.title = props.title;
  }
  if (Object.keys(pictureNames).length > 0) {
    image.pictureNames = pictureNames;
  }
  if (props.decorative !== undefined) {
    image.decorative = props.decorative;
  }
  if (props.hidden !== undefined) {
    image.hidden = props.hidden;
  }
  if (props.docPrExtensions !== undefined) {
    image.docPrExtensions = props.docPrExtensions;
  }
  const safeSrc = sanitizeImageSrc(imageData.src);
  if (safeSrc) {
    image.src = safeSrc;
  }
  if (imageData.mimeType) {
    image.mimeType = imageData.mimeType;
  }
  if (imageData.filename) {
    image.filename = imageData.filename;
  }
  if (padding) {
    image.padding = padding;
  }
  if (transform) {
    image.transform = transform;
  }
  if (crop) {
    image.crop = crop;
  }
  if (opacity !== undefined) {
    image.opacity = opacity;
  }
  if (effects !== undefined) {
    image.effects = effects;
  }
  if (frameLocks) {
    image.frameLocks = frameLocks;
  }

  // The `wp:docPr` links, target-checked. Mirrors hyperlinkParser.ts: an
  // unsafe or unresolved target leaves the link off the image rather than
  // storing a raw javascript:/data:/file: href.
  Object.assign(image, safeDocPrLinks({ props, rels }));

  return image;
}

/**
 * Parse a wp:anchor element (floating/anchored image)
 *
 * @param anchorEl - The wp:anchor element
 * @param rels - Relationship map for resolving rId
 * @param media - Media files map
 * @returns Parsed Image object
 */
function parseAnchor(
  anchorEl: XmlElement,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined,
): Image {
  // Parse extent (size)
  const extent = findDrawingChild(anchorEl, "extent");
  const size = parseExtent(extent);

  // Parse effect extent
  const effectExtent = findDrawingChild(anchorEl, "effectExtent");
  const padding = parseDrawingEffectExtent(effectExtent);

  // Parse document properties
  const docPr = findDrawingChild(anchorEl, "docPr");
  const props = parseDocProps(docPr);

  const frameLocks = parseGraphicFrameLocks(anchorEl);

  // `pic:cNvPr`'s own set, kept apart from `wp:docPr`'s above: the rebuild
  // wrote the media filename here, renaming every picture the author had named.
  const pictureNames = parseNonVisualDrawingNames(findPictureNonVisualProps(anchorEl));

  const behindDoc = parseAnchorBehindDoc(anchorEl);

  // `CT_Anchor`'s own attributes and its `wp:simplePos`, from the one owner an
  // image, a shape and a text box share: absent states nothing, and a rebuild
  // writes OOXML's default rather than a constant folio chose.
  const anchor = parseDrawingAnchor(anchorEl);

  // Read distance attributes from the wp:anchor element itself (fallback values)
  const anchorDistT = parseNumericAttribute(anchorEl, null, "distT");
  const anchorDistB = parseNumericAttribute(anchorEl, null, "distB");
  const anchorDistL = parseNumericAttribute(anchorEl, null, "distL");
  const anchorDistR = parseNumericAttribute(anchorEl, null, "distR");
  const anchorDistances = {
    ...(anchorDistT != null ? { distT: anchorDistT } : {}),
    ...(anchorDistB != null ? { distB: anchorDistB } : {}),
    ...(anchorDistL != null ? { distL: anchorDistL } : {}),
    ...(anchorDistR != null ? { distR: anchorDistR } : {}),
  };

  // Parse wrap element (wrap child values take priority over anchor-level values)
  const anchorEffectExtent = parseDrawingEffectExtent(effectExtent);
  const wrap = parseWrapElement({
    wrapEl: findWrapElement(anchorEl),
    behindDoc,
    anchorDistances,
    ...(anchorEffectExtent === undefined ? {} : { anchorEffectExtent }),
  });

  // Parse position
  const posH = findDrawingChild(anchorEl, "positionH");
  const posV = findDrawingChild(anchorEl, "positionV");
  const horizontal = parsePositionH(posH);
  const vertical = parsePositionV(posV);

  let position: ImagePosition | undefined;
  if (horizontal || vertical) {
    position = {
      horizontal: horizontal ?? { relativeTo: "column" },
      vertical: vertical ?? { relativeTo: "paragraph" },
    };
  }

  // Find blip and extract rId
  const blipFill = findBlipFillElement(anchorEl);
  const blip = blipFill ? findByFullName(blipFill, "a:blip") : null;
  const rId = extractBlipRId(blip);
  const crop = parseImageCrop(blipFill);
  const opacity = parseImageOpacity(blip);
  const effects = parseImageLuminance(blip);

  // Resolve image data
  const imageData = resolveImageData(rId, rels, media);

  // Find transform
  const xfrm = findPictureTransform(anchorEl);
  const transform = parseTransform(xfrm);

  const image: Image = {
    type: "image",
    ...(rId === undefined ? {} : { rId }),
    size,
    wrap,
  };

  // Add optional properties
  if (props.id) {
    image.id = props.id;
  }
  if (props.name !== undefined) {
    image.docPrName = props.name;
  }
  if (props.alt !== undefined) {
    image.alt = props.alt;
  }
  if (props.title !== undefined) {
    image.title = props.title;
  }
  if (Object.keys(pictureNames).length > 0) {
    image.pictureNames = pictureNames;
  }
  if (props.decorative !== undefined) {
    image.decorative = props.decorative;
  }
  if (props.hidden !== undefined) {
    image.hidden = props.hidden;
  }
  if (props.docPrExtensions !== undefined) {
    image.docPrExtensions = props.docPrExtensions;
  }
  const safeSrc = sanitizeImageSrc(imageData.src);
  if (safeSrc) {
    image.src = safeSrc;
  }
  if (imageData.mimeType) {
    image.mimeType = imageData.mimeType;
  }
  if (imageData.filename) {
    image.filename = imageData.filename;
  }
  if (position) {
    image.position = position;
  }
  if (padding) {
    image.padding = padding;
  }
  if (transform) {
    image.transform = transform;
  }
  if (crop) {
    image.crop = crop;
  }
  if (opacity !== undefined) {
    image.opacity = opacity;
  }
  if (effects !== undefined) {
    image.effects = effects;
  }
  if (frameLocks) {
    image.frameLocks = frameLocks;
  }
  if (anchor !== undefined) {
    image.anchor = anchor;
  }

  // The `wp:docPr` links, target-checked. Mirrors hyperlinkParser.ts: an
  // unsafe or unresolved target leaves the link off the image rather than
  // storing a raw javascript:/data:/file: href.
  Object.assign(image, safeDocPrLinks({ props, rels }));

  return image;
}

/**
 * Parse a w:drawing element
 *
 * The drawing element contains either wp:inline or wp:anchor.
 *
 * @param drawingEl - The w:drawing element
 * @param rels - Relationship map for resolving rId
 * @param media - Media files map
 * @returns Parsed Image object or null if not an image
 */
export function parseDrawing(
  drawingEl: XmlElement,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined,
): Image | null {
  // Skip text box shapes — they are handled by textBoxParser, not as images
  if (isTextBoxDrawing(drawingEl)) {
    return null;
  }

  // Which of the two a drawing carries is its anchoring, and the namespace is
  // what says so. A producer binding WordprocessingDrawing to a prefix other
  // than `wp` writes the same document, and matching the spelling read it as
  // no drawing at all.
  const inline = findDrawingChild(drawingEl, "inline");
  if (inline) {
    return parseInline(inline, rels, media);
  }
  const anchor = findDrawingChild(drawingEl, "anchor");
  return anchor ? parseAnchor(anchor, rels, media) : null;
}

/**
 * Parse an image from a w:drawing element
 *
 * This is the main entry point for image parsing.
 *
 * @param node - The w:drawing XML element
 * @param rels - Relationship map for resolving rId
 * @param media - Media files map
 * @returns Parsed Image object or null if parsing fails
 */
export function parseImage(
  node: XmlElement,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined,
): Image | null {
  return parseDrawing(node, rels, media);
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Check if an image is inline (not floating)
 */
export function isInlineImage(image: Image): boolean {
  return image.wrap.type === "inline";
}

/**
 * Check if an image is floating (anchored)
 */
export function isFloatingImage(image: Image): boolean {
  return image.wrap.type !== "inline";
}

/**
 * Check if an image is behind text
 */
export function isBehindText(image: Image): boolean {
  return image.wrap.type === "behind";
}

/**
 * Check if an image is in front of text
 */
export function isInFrontOfText(image: Image): boolean {
  return image.wrap.type === "inFront";
}

/**
 * Get image width in pixels
 */
export function getImageWidthPx(image: Image): number {
  return emuToPixels(image.size.width);
}

/**
 * Get image height in pixels
 */
export function getImageHeightPx(image: Image): number {
  return emuToPixels(image.size.height);
}

/**
 * Get image dimensions in pixels
 */
export function getImageDimensionsPx(image: Image): {
  width: number;
  height: number;
} {
  return {
    width: emuToPixels(image.size.width),
    height: emuToPixels(image.size.height),
  };
}

/**
 * Check if image has alt text (for accessibility)
 */
export function hasAltText(image: Image): boolean {
  return !!image.alt && image.alt.trim().length > 0;
}

/**
 * Check if image is decorative (should be ignored by screen readers)
 */
export function isDecorativeImage(image: Image): boolean {
  return image.decorative === true;
}

/**
 * Get wrap distances in pixels
 */
export function getWrapDistancesPx(image: Image): {
  top: number;
  bottom: number;
  left: number;
  right: number;
} {
  return {
    top: emuToPixels(image.wrap.distT),
    bottom: emuToPixels(image.wrap.distB),
    left: emuToPixels(image.wrap.distL),
    right: emuToPixels(image.wrap.distR),
  };
}

/**
 * Check if image needs text wrapping
 */
export function needsTextWrapping(image: Image): boolean {
  const wrapTypes = ["square", "tight", "through", "topAndBottom"];
  return wrapTypes.includes(image.wrap.type);
}
