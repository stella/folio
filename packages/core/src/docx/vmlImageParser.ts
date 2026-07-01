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
import { elementToXml, findAllDeep, findChild, getAttribute, type XmlElement } from "./xmlParser";

/** Convert a CSS length (pt/in/px/cm/mm/pc, default px) to pixels. */
function cssLengthToPx(raw: string | undefined): number | undefined {
  if (!raw) {
    return undefined;
  }
  const match = /^(?<amount>-?[\d.]+)\s*(?<unit>pt|in|px|cm|mm|pc)?$/u.exec(raw.trim());
  const amountText = match?.groups?.["amount"];
  if (amountText === undefined) {
    return undefined;
  }
  const value = Number.parseFloat(amountText);
  if (Number.isNaN(value)) {
    return undefined;
  }
  switch (match?.groups?.["unit"]) {
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

/** Read the `r:id` (or fallbacks) off a `v:imagedata` element. */
function readImageDataRId(imagedata: XmlElement): string {
  return (
    getAttribute(imagedata, "r", "id") ??
    getAttribute(imagedata, "r", "embed") ??
    getAttribute(imagedata, null, "id") ??
    ""
  );
}

/**
 * Parse a `w:pict` element into an inline image, or null when it carries no
 * ordinary VML picture (no resolvable `<v:imagedata>`, or a watermark shape).
 */
export function parseVmlImageContent(
  pictElement: XmlElement,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
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
    // synthesized DrawingML `<w:drawing>`.
    return { type: "drawing", image, rawXml: elementToXml(pictElement) };
  }

  return null;
}
