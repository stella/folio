import type { DrawingContent } from "../model/document";
import { MAX_TABLE_PICTURE_ID } from "../model/formatting";
import {
  rewriteStreamingXmlDecimalAttributes,
  scanStreamingXmlElements,
} from "../xml/streamingXmlParser";
import type { XmlElement } from "../xml/xmlElement";
import { OOXML_NS, resolveNamespaceUri } from "../xml/xmlNamespaceContext";

const STRICT_DRAWING = "http://purl.oclc.org/ooxml/drawingml/";

// Captures can omit the namespaces a standard story supplies. Local
// declarations inside the capture still take precedence.
const DRAWING_FRAGMENT_START = `<part xmlns:w="${OOXML_NS.w}" xmlns:wp="${OOXML_NS.wp}" xmlns:pic="${OOXML_NS.pic}" xmlns:a="${OOXML_NS.a}" xmlns:r="${OOXML_NS.r}">`;
const DRAWING_FRAGMENT_END = "</part>";
const drawingFragment = (xml: string) =>
  `${DRAWING_FRAGMENT_START}${xml}${DRAWING_FRAGMENT_END}`;

const namespaceKind = (uri: string | undefined) => {
  switch (uri) {
    case OOXML_NS.w:
    case "http://purl.oclc.org/ooxml/wordprocessingml/main":
      return "word";
    case OOXML_NS.wp:
    case `${STRICT_DRAWING}wordprocessingDrawing`:
      return "anchor";
    case OOXML_NS.pic:
    case `${STRICT_DRAWING}picture`:
      return "picture";
    case OOXML_NS.a:
    case `${STRICT_DRAWING}main`:
      return "drawing";
    default:
      return undefined;
  }
};

const localName = ({ name }: XmlElement) => name?.slice(name.indexOf(":") + 1);

const PICTURE_ELEMENTS = {
  word: new Set(["drawing"]),
  anchor: new Set([
    "inline",
    "extent",
    "effectExtent",
    "docPr",
    "cNvGraphicFramePr",
  ]),
  picture: new Set(["pic", "nvPicPr", "cNvPr", "cNvPicPr", "blipFill", "spPr"]),
  drawing: new Set([
    "graphic",
    "graphicData",
    "graphicFrameLocks",
    "picLocks",
    "blip",
    "srcRect",
    "stretch",
    "fillRect",
    "xfrm",
    "off",
    "ext",
    "prstGeom",
    "avLst",
    "alphaModFix",
  ]),
};

const numericDrawingId = (value: unknown): number | undefined => {
  if (typeof value !== "string" || !/^[1-9]\d*$/u.test(value)) return undefined;
  const id = Number(value);
  return id <= MAX_TABLE_PICTURE_ID ? id : undefined;
};

const isDocPr = (element: XmlElement) =>
  namespaceKind(element.namespaceUri) === "anchor" &&
  localName(element) === "docPr";

/** Every authored drawing identity in a captured XML fragment. */
export const readDrawingIds = (xml: string, ids: Set<number>) => {
  scanStreamingXmlElements(drawingFragment(xml), ({ element }) => {
    if (!isDocPr(element)) return;
    const id = numericDrawingId(element.attributes?.["id"]);
    if (id !== undefined) ids.add(id);
  });
};

/** The move halves can share a picture relationship in the same part. Each
 * inline DrawingML object still needs its own docPr ID while pending. */
export const movableTablePictureId = (drawing: DrawingContent) => {
  const id = numericDrawingId(drawing.image.id);
  if (
    id === undefined ||
    drawing.rawXmlMode !== undefined ||
    drawing.image.wrap.type !== "inline" ||
    drawing.image.rId === undefined ||
    drawing.image.hlinkRId !== undefined ||
    drawing.image.hlinkHoverXml !== undefined ||
    drawing.image.docPrExtensions !== undefined
  )
    return undefined;
  if (drawing.rawXml === undefined) return id;

  let docPr = 0;
  let picture = 0;
  let images = 0;
  let portable = true;
  const scanned = scanStreamingXmlElements(
    drawingFragment(drawing.rawXml),
    ({ element, parent }) => {
      if (parent === undefined || !portable) return;
      const kind = namespaceKind(element.namespaceUri);
      const name = localName(element);
      if (
        kind === undefined ||
        name === undefined ||
        !PICTURE_ELEMENTS[kind].has(name)
      ) {
        portable = false;
        return;
      }
      for (const [attribute, value] of Object.entries(
        element.attributes ?? {},
      )) {
        const colon = attribute.indexOf(":");
        if (colon < 0 || attribute.startsWith("xmlns:")) continue;
        const uri = resolveNamespaceUri(
          element.namespaceScope,
          attribute.slice(0, colon),
        );
        if (
          kind !== "drawing" ||
          name !== "blip" ||
          attribute.slice(colon + 1) !== "embed" ||
          (uri !== OOXML_NS.r &&
            uri !==
              "http://purl.oclc.org/ooxml/officeDocument/relationships") ||
          value !== drawing.image.rId
        ) {
          portable = false;
          return;
        }
        images += 1;
      }
      if (kind === "anchor" && name === "docPr") {
        docPr += 1;
        if (element.attributes?.["id"] !== String(id)) portable = false;
      } else if (kind === "picture" && name === "cNvPr") {
        picture += 1;
        const pictureId = element.attributes?.["id"];
        if (pictureId !== "0" && pictureId !== String(id)) portable = false;
      }
    },
  );
  return scanned.status === "scanned" &&
    portable &&
    docPr === 1 &&
    picture === 1 &&
    images === 1
    ? id
    : undefined;
};

type PictureIdentityChange = { source: number; destination: number };

/** Rewrite only the drawing identifiers, preserving the captured picture bytes. */
export const withTablePictureIdentity = (
  drawing: DrawingContent,
  { source, destination }: PictureIdentityChange,
): DrawingContent => {
  if (drawing.rawXmlMode !== undefined) return drawing;
  const image = { ...drawing.image, id: String(destination) };
  if (drawing.rawXml === undefined) return { ...drawing, image };
  const rewrite = rewriteStreamingXmlDecimalAttributes(
    drawingFragment(drawing.rawXml),
    (element) => {
      const picture =
        namespaceKind(element.namespaceUri) === "picture" &&
        localName(element) === "cNvPr";
      return (picture || isDocPr(element)) &&
        element.attributes?.["id"] === String(source)
        ? new Map([["id", String(destination)]])
        : null;
    },
  );
  return rewrite.status === "rewritten"
    ? {
        ...drawing,
        image,
        rawXml: rewrite.value.slice(
          DRAWING_FRAGMENT_START.length,
          -DRAWING_FRAGMENT_END.length,
        ),
      }
    : drawing;
};
