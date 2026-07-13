/**
 * VML image parser — legacy `w:pict` inline pictures.
 *
 * Older Word documents (and some exporters) embed inline pictures such as
 * header logos as VML instead of DrawingML:
 *
 *   <w:r><w:pict>
 *     <v:shape id="Picture 1" type="#_x0000_t75" style="width:120pt;height:40pt">
 *       <v:imagedata r:id="rId7" o:title="logo"/>
 *     </v:shape>
 *   </w:pict></w:r>
 *
 * The run parser used to drop `w:pict` entirely, so these pictures never
 * rendered. This resolves `<v:imagedata r:id>` to the same media part a
 * DrawingML image uses and produces the identical `DrawingContent` / `Image`
 * node, so the rest of the pipeline (ProseMirror conversion, layout, painter)
 * renders it through the existing inline-image path.
 *
 * The original VML is preserved verbatim on the returned drawing's `rawXml`:
 * the serializer has no VML synthesis path (it only emits DrawingML), so
 * without this a `w:pict` would round-trip as a `<w:drawing>`. Emitting the
 * captured XML keeps save byte-for-byte for these runs.
 *
 * Picture watermarks (`WordPictureWatermark…` shapes) are owned by
 * watermarkParser and stripped before body parsing; the `isWatermarkShape`
 * guard here is defensive so such a shape is never rendered twice.
 *
 * Ported from eigenpal/docx-editor `vmlImageParser.ts`.
 */

import type { DrawingContent, Image, MediaFile, RelationshipMap } from "../types/document";
import { pixelsToEmu } from "../utils/units";
import { resolveImageData } from "./imageParser";
import { isWatermarkShape } from "./watermarkParser";
import {
  cloneWithXmlnsDeclarations,
  elementToXml,
  findAllDeep,
  findChild,
  getChildElements,
  getAttribute,
  getLocalName,
  type XmlElement,
} from "./xmlParser";

const MAX_VML_PREVIEW_SHAPES = 256;
const MAX_VML_PREVIEW_DIMENSION_PX = 20_000;
const MAX_VML_SVG_CHARACTERS = 1_000_000;
const SAFE_VML_COLORS = new Set([
  "black",
  "white",
  "red",
  "green",
  "blue",
  "yellow",
  "gray",
  "grey",
  "silver",
  "maroon",
  "purple",
  "fuchsia",
  "lime",
  "olive",
  "navy",
  "teal",
  "aqua",
]);

/**
 * Convert a CSS length (pt/in/px/cm/mm/pc, default px) to pixels. Units are
 * matched case-insensitively (`2IN`, `2Pt` are valid VML). The numeric part
 * must be a single well-formed decimal, so malformed input like `1.2.3` is
 * rejected rather than truncated to `1.2`.
 */
function cssLengthToPx(raw: string | undefined): number | undefined {
  if (!raw) {
    return undefined;
  }
  const match = /^(?<amount>-?(?:\d+(?:\.\d+)?|\.\d+))\s*(?<unit>pt|in|px|cm|mm|pc)?$/iu.exec(
    raw.trim(),
  );
  const amountText = match?.groups?.["amount"];
  if (amountText === undefined) {
    return undefined;
  }
  const value = Number.parseFloat(amountText);
  if (Number.isNaN(value)) {
    return undefined;
  }
  switch (match?.groups?.["unit"]?.toLowerCase()) {
    case "pt":
      return (value / 72) * 96;
    case "in":
      return value * 96;
    case "cm":
      return (value / 2.54) * 96;
    case "mm":
      return (value / 25.4) * 96;
    case "pc":
      return (value / 6) * 96;
    case "px":
    case undefined:
      return value;
    default:
      return undefined;
  }
}

/** Read a `style="k1:v1;k2:v2"` attribute into a lowercased key/value record. */
function parseStyleAttr(style: string | null): Record<string, string> {
  const out: Record<string, string> = {};
  if (!style) {
    return out;
  }
  for (const decl of style.split(";")) {
    const colon = decl.indexOf(":");
    if (colon < 0) {
      continue;
    }
    const key = decl.slice(0, colon).trim().toLowerCase();
    const value = decl.slice(colon + 1).trim();
    if (key) {
      out[key] = value;
    }
  }
  return out;
}

const finiteNumber = (raw: string | null | undefined): number | undefined => {
  if (!raw || !/^-?(?:\d+(?:\.\d+)?|\.\d+)$/u.test(raw.trim())) {
    return undefined;
  }
  const value = Number.parseFloat(raw);
  return Number.isFinite(value) ? value : undefined;
};

const coordinatePair = (raw: string | null): { first: number; second: number } | undefined => {
  const [firstRaw, secondRaw, extra] = raw?.split(",").map((part) => part.trim()) ?? [];
  if (extra !== undefined) {
    return undefined;
  }
  const first = finiteNumber(firstRaw);
  const second = finiteNumber(secondRaw);
  return first === undefined || second === undefined ? undefined : { first, second };
};

const safeColor = (raw: string | null, fallback: string): string => {
  const normalized = raw?.trim().toLowerCase();
  if (!normalized) {
    return fallback;
  }
  if (SAFE_VML_COLORS.has(normalized)) {
    return normalized;
  }
  if (/^#[0-9a-f]{3}(?:[0-9a-f]{3})?$/u.test(normalized)) {
    return normalized;
  }
  return /^[0-9a-f]{6}$/u.test(normalized) ? `#${normalized}` : fallback;
};

const validPreviewDimension = (value: number | undefined): value is number =>
  value !== undefined && value > 0 && value <= MAX_VML_PREVIEW_DIMENSION_PX;

const svgDataUrl = (svg: string): string | undefined =>
  svg.length <= MAX_VML_SVG_CHARACTERS
    ? `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`
    : undefined;

const vmlShapeSvg = (shape: XmlElement, width: number, height: number, x = 0, y = 0): string => {
  const fill = safeColor(getAttribute(shape, null, "fillcolor"), "white");
  const stroked = getAttribute(shape, null, "stroked")?.toLowerCase() !== "f";
  const stroke = stroked ? safeColor(getAttribute(shape, null, "strokecolor"), "black") : "none";
  const localName = getLocalName(shape.name ?? "");
  if (localName === "oval") {
    return `<ellipse cx="${x + width / 2}" cy="${y + height / 2}" rx="${width / 2}" ry="${height / 2}" fill="${fill}" stroke="${stroke}"/>`;
  }
  const radius = localName === "roundrect" ? Math.min(width, height) * 0.1 : 0;
  return `<rect x="${x}" y="${y}" width="${width}" height="${height}" rx="${radius}" fill="${fill}" stroke="${stroke}"/>`;
};

const previewImage = (
  pictElement: XmlElement,
  svg: string,
  widthPx: number,
  heightPx: number,
  style: Record<string, string>,
  rootXmlns: Record<string, string>,
): DrawingContent | null => {
  const src = svgDataUrl(svg);
  if (!src) {
    return null;
  }
  const leftPx = cssLengthToPx(style["margin-left"] ?? style["left"]) ?? 0;
  const topPx = cssLengthToPx(style["margin-top"] ?? style["top"]) ?? 0;
  const horizontalRelative = style["mso-position-horizontal-relative"] === "page";
  const verticalRelative = style["mso-position-vertical-relative"] === "page";
  const zIndex = finiteNumber(style["z-index"]);
  const image: Image = {
    type: "image",
    rId: "",
    src,
    mimeType: "image/svg+xml",
    filename: "vml-shape-preview.svg",
    size: { width: pixelsToEmu(widthPx), height: pixelsToEmu(heightPx) },
    wrap: { type: zIndex !== undefined && zIndex >= 0 ? "inFront" : "behind" },
    position: {
      horizontal: {
        relativeTo: horizontalRelative ? "page" : "character",
        posOffset: pixelsToEmu(leftPx),
      },
      vertical: {
        relativeTo: verticalRelative ? "page" : "paragraph",
        posOffset: pixelsToEmu(topPx),
      },
    },
  };
  return {
    type: "drawing",
    image,
    rawXml: elementToXml(cloneWithXmlnsDeclarations(pictElement, rootXmlns)),
  };
};

const parseStandaloneShapePreview = (
  pictElement: XmlElement,
  shape: XmlElement,
  rootXmlns: Record<string, string>,
): DrawingContent | null => {
  const style = parseStyleAttr(getAttribute(shape, null, "style"));
  const widthPx = cssLengthToPx(style["width"]);
  const heightPx = cssLengthToPx(style["height"]);
  if (!validPreviewDimension(widthPx) || !validPreviewDimension(heightPx)) {
    return null;
  }
  const content = vmlShapeSvg(shape, widthPx, heightPx);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${widthPx} ${heightPx}" width="${widthPx}" height="${heightPx}">${content}</svg>`;
  return previewImage(pictElement, svg, widthPx, heightPx, style, rootXmlns);
};

const parseGroupPreview = (
  pictElement: XmlElement,
  group: XmlElement,
  rootXmlns: Record<string, string>,
): DrawingContent | null => {
  const style = parseStyleAttr(getAttribute(group, null, "style"));
  const widthPx = cssLengthToPx(style["width"]);
  const heightPx = cssLengthToPx(style["height"]);
  if (!validPreviewDimension(widthPx) || !validPreviewDimension(heightPx)) {
    return null;
  }
  const origin = coordinatePair(getAttribute(group, null, "coordorigin")) ?? {
    first: 0,
    second: 0,
  };
  const size = coordinatePair(getAttribute(group, null, "coordsize")) ?? {
    first: widthPx,
    second: heightPx,
  };
  if (size.first <= 0 || size.second <= 0) {
    return null;
  }
  const content = getChildElements(group)
    .slice(0, MAX_VML_PREVIEW_SHAPES)
    .map((child) => {
      const localName = getLocalName(child.name ?? "");
      if (localName === "line") {
        const from = coordinatePair(getAttribute(child, null, "from"));
        const to = coordinatePair(getAttribute(child, null, "to"));
        if (!from || !to) {
          return "";
        }
        const stroke = safeColor(getAttribute(child, null, "strokecolor"), "black");
        return `<line x1="${from.first}" y1="${from.second}" x2="${to.first}" y2="${to.second}" stroke="${stroke}"/>`;
      }
      if (localName !== "rect" && localName !== "roundrect" && localName !== "oval") {
        return "";
      }
      const childStyle = parseStyleAttr(getAttribute(child, null, "style"));
      const x = finiteNumber(childStyle["left"]);
      const y = finiteNumber(childStyle["top"]);
      const width = finiteNumber(childStyle["width"]);
      const height = finiteNumber(childStyle["height"]);
      if (x === undefined || y === undefined || width === undefined || height === undefined) {
        return "";
      }
      return vmlShapeSvg(child, width, height, x, y);
    })
    .join("");
  if (!content) {
    return null;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${origin.first} ${origin.second} ${size.first} ${size.second}" width="${widthPx}" height="${heightPx}">${content}</svg>`;
  return previewImage(pictElement, svg, widthPx, heightPx, style, rootXmlns);
};

/**
 * Read the relationship id off a `v:imagedata` element. Word writes `r:id`;
 * some legacy / third-party generators use `r:embed` or the office-namespace
 * `o:relid` instead, so fall back through those before the bare `id`.
 */
function readImageDataRId(imagedata: XmlElement): string {
  return (
    getAttribute(imagedata, "r", "id") ??
    getAttribute(imagedata, "r", "embed") ??
    getAttribute(imagedata, "o", "relid") ??
    getAttribute(imagedata, null, "id") ??
    ""
  );
}

/**
 * Parse a `w:pict` element into an inline image, or null when it carries no
 * ordinary VML picture (no resolvable `<v:imagedata>`, or a watermark shape).
 *
 * `rootXmlns` carries the source document/header namespace declarations. They
 * are injected onto the captured VML so the raw replay stays self-contained
 * when the producer bound WordprocessingML / relationship / VML namespaces to
 * non-canonical prefixes the serializer's root does not declare.
 */
export function parseVmlImageContent(
  pictElement: XmlElement,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  rootXmlns: Record<string, string> = {},
): DrawingContent | null {
  // A VML picture's image lives in <v:imagedata r:id> inside a shape
  // (v:shape / v:rect / v:roundrect / v:oval). Walk each shape kind and look
  // for an imagedata child within it.
  const shapes = [
    ...findAllDeep(pictElement, "v", "shape"),
    ...findAllDeep(pictElement, "v", "rect"),
    ...findAllDeep(pictElement, "v", "roundrect"),
    ...findAllDeep(pictElement, "v", "oval"),
  ];

  for (const shape of shapes) {
    const imagedata = findChild(shape, "v", "imagedata");
    if (!imagedata) {
      continue;
    }

    const rId = readImageDataRId(imagedata);
    if (!rId) {
      continue;
    }

    // Watermarks are owned by watermarkParser; rendering them here too would
    // duplicate them.
    if (isWatermarkShape(shape)) {
      continue;
    }

    const { src, mimeType, filename } = resolveImageData(
      rId,
      rels ?? undefined,
      media ?? undefined,
    );

    const shapeStyle = parseStyleAttr(getAttribute(shape, null, "style"));
    const widthPx = cssLengthToPx(shapeStyle["width"]);
    const heightPx = cssLengthToPx(shapeStyle["height"]);

    // VML pictures in a run are inline-flow. Absolute-positioned VML is treated
    // as inline here too: rendering the picture in flow is far better than
    // dropping it, and the original XML round-trips verbatim via rawXml.
    const image: Image = {
      type: "image",
      rId,
      size: {
        width: widthPx != null ? pixelsToEmu(widthPx) : 0,
        height: heightPx != null ? pixelsToEmu(heightPx) : 0,
      },
      wrap: { type: "inline" },
    };
    if (src) {
      image.src = src;
    }
    if (mimeType) {
      image.mimeType = mimeType;
    }
    if (filename) {
      image.filename = filename;
    }
    const title = getAttribute(imagedata, "o", "title");
    if (title) {
      image.title = title;
    }

    // Preserve the exact VML so the serializer replays it instead of emitting a
    // synthesized DrawingML `<w:drawing>`. Inject the source root's namespace
    // declarations so a non-canonical prefix still resolves in the replay.
    return {
      type: "drawing",
      image,
      rawXml: elementToXml(cloneWithXmlnsDeclarations(pictElement, rootXmlns)),
    };
  }

  for (const child of getChildElements(pictElement)) {
    const localName = getLocalName(child.name ?? "");
    if (localName === "group") {
      const preview = parseGroupPreview(pictElement, child, rootXmlns);
      if (preview) {
        return preview;
      }
      continue;
    }
    if (localName === "rect" || localName === "roundrect" || localName === "oval") {
      const preview = parseStandaloneShapePreview(pictElement, child, rootXmlns);
      if (preview) {
        return preview;
      }
    }
  }

  return null;
}
