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

import type {
  DrawingContent,
  Image,
  ImagePosition,
  MediaFile,
  RelationshipMap,
} from "../types/document";
import { sanitizeImageSrc } from "../utils/sanitizeImageSrc";
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
import {
  isValidVmlPreviewDimension,
  parseVmlNumber,
  parseVmlStyle,
  renderStandaloneVmlPreview,
  renderVmlGroupPreview,
  vmlCssLengthToPx,
  vmlSvgDataUrl,
  type VmlPreviewResult,
} from "./vmlPreview";

const VML_POSITION_ABSOLUTE = "absolute";
const IMAGE_WRAP_INLINE = "inline";
const IMAGE_WRAP_BEHIND = "behind";
const IMAGE_WRAP_IN_FRONT = "inFront";

const pixelsToSafeEmu = (pixels: number): number | undefined => {
  const emu = pixelsToEmu(pixels);
  return Number.isSafeInteger(emu) ? emu : undefined;
};

const horizontalRelativeTo = (
  value: string | undefined,
): ImagePosition["horizontal"]["relativeTo"] => {
  switch (value?.toLowerCase()) {
    case "char":
      return "character";
    case "text":
      return "column";
    case "margin":
      return "margin";
    case "page":
      return "page";
    case "left-margin-area":
      return "leftMargin";
    case "right-margin-area":
      return "rightMargin";
    case "inner-margin-area":
      return "insideMargin";
    case "outer-margin-area":
      return "outsideMargin";
    default:
      return "character";
  }
};

const verticalRelativeTo = (value: string | undefined): ImagePosition["vertical"]["relativeTo"] => {
  switch (value?.toLowerCase()) {
    case "line":
      return "line";
    case "text":
      return "paragraph";
    case "margin":
      return "margin";
    case "page":
      return "page";
    case "top-margin-area":
      return "topMargin";
    case "bottom-margin-area":
      return "bottomMargin";
    case "inner-margin-area":
      return "insideMargin";
    case "outer-margin-area":
      return "outsideMargin";
    default:
      return "paragraph";
  }
};

const vmlImageLayout = (
  style: Record<string, string>,
): Pick<Image, "wrap"> & { position?: ImagePosition } => {
  if (style["position"]?.toLowerCase() !== VML_POSITION_ABSOLUTE) {
    return { wrap: { type: IMAGE_WRAP_INLINE } };
  }

  const leftPx = vmlCssLengthToPx(style["margin-left"] ?? style["left"]) ?? 0;
  const topPx = vmlCssLengthToPx(style["margin-top"] ?? style["top"]) ?? 0;
  const zIndex = parseVmlNumber(style["z-index"]);
  return {
    wrap: {
      type: zIndex !== undefined && zIndex < 0 ? IMAGE_WRAP_BEHIND : IMAGE_WRAP_IN_FRONT,
    },
    position: {
      horizontal: {
        relativeTo: horizontalRelativeTo(style["mso-position-horizontal-relative"]),
        posOffset: pixelsToSafeEmu(leftPx) ?? 0,
      },
      vertical: {
        relativeTo: verticalRelativeTo(style["mso-position-vertical-relative"]),
        posOffset: pixelsToSafeEmu(topPx) ?? 0,
      },
    },
  };
};

const previewImage = (
  pictElement: XmlElement,
  svg: string,
  widthPx: number,
  heightPx: number,
  style: Record<string, string>,
  rootXmlns: Record<string, string>,
): DrawingContent | null => {
  const src = vmlSvgDataUrl(svg);
  if (!src) {
    return null;
  }
  const leftPx = vmlCssLengthToPx(style["margin-left"] ?? style["left"]) ?? 0;
  const topPx = vmlCssLengthToPx(style["margin-top"] ?? style["top"]) ?? 0;
  const horizontalRelative = style["mso-position-horizontal-relative"] === "page";
  const verticalRelative = style["mso-position-vertical-relative"] === "page";
  const zIndex = parseVmlNumber(style["z-index"]);
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
        posOffset: pixelsToSafeEmu(leftPx) ?? 0,
      },
      vertical: {
        relativeTo: verticalRelative ? "page" : "paragraph",
        posOffset: pixelsToSafeEmu(topPx) ?? 0,
      },
    },
  };
  return {
    type: "drawing",
    image,
    rawXml: elementToXml(cloneWithXmlnsDeclarations(pictElement, rootXmlns)),
  };
};

const previewDrawing = (
  pictElement: XmlElement,
  preview: VmlPreviewResult,
  rootXmlns: Record<string, string>,
): DrawingContent | null => {
  switch (preview.type) {
    case "rendered":
      return previewImage(
        pictElement,
        preview.svg,
        preview.widthPx,
        preview.heightPx,
        preview.style,
        rootXmlns,
      );
    case "empty":
    case "invalid":
      return null;
    default:
      return preview satisfies never;
  }
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

    const shapeStyle = parseVmlStyle(getAttribute(shape, null, "style"));
    const widthPx = vmlCssLengthToPx(shapeStyle["width"]);
    const heightPx = vmlCssLengthToPx(shapeStyle["height"]);
    const isEmbeddedObject = getLocalName(pictElement.name ?? "") === "object";
    if (
      isEmbeddedObject &&
      ((widthPx !== undefined && !isValidVmlPreviewDimension(widthPx)) ||
        (heightPx !== undefined && !isValidVmlPreviewDimension(heightPx)))
    ) {
      continue;
    }

    // Static VML pictures participate in the run's inline flow. Absolute VML
    // pictures are page artwork: preserve their authored anchor so they paint
    // at the correct location without contributing to paragraph height.
    const image: Image = {
      type: "image",
      rId,
      size: {
        width: widthPx != null ? (pixelsToSafeEmu(widthPx) ?? 0) : 0,
        height: heightPx != null ? (pixelsToSafeEmu(heightPx) ?? 0) : 0,
      },
      ...vmlImageLayout(shapeStyle),
    };
    const safeSrc = sanitizeImageSrc(src);
    if (safeSrc) {
      image.src = safeSrc;
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
      const preview = previewDrawing(pictElement, renderVmlGroupPreview(child), rootXmlns);
      if (preview) {
        return preview;
      }
      continue;
    }
    if (localName === "rect" || localName === "roundrect" || localName === "oval") {
      const preview = previewDrawing(pictElement, renderStandaloneVmlPreview(child), rootXmlns);
      if (preview) {
        return preview;
      }
    }
  }

  return null;
}
