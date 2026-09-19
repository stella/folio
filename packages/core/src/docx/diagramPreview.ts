import type { Image, MediaFile, RelationshipMap } from "../types/document";
import {
  findChildByNamespaceUri,
  getAttribute,
  getAttributeByNamespaceUri,
  getLocalName,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  parseNumericAttribute,
  parseXmlDocument,
} from "./xmlParser";
import { resolveRelationshipIdOfType, resolveRelativePath } from "./relsParser";
import type { XmlElement } from "./xmlParser";

const MAX_PREVIEW_SHAPES = 128;
const MAX_PREVIEW_PIXELS = 1_440_000;
const MAX_PREVIEW_PAINT_PIXELS = MAX_PREVIEW_PIXELS * 4;
const DRAWINGML_NAMESPACE_URIS = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/main",
  "http://purl.oclc.org/ooxml/drawingml/main",
]);
const DIAGRAM_NAMESPACE_URIS = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/diagram",
  "http://purl.oclc.org/ooxml/drawingml/diagram",
]);
const DIAGRAM_DRAWING_NAMESPACE_URIS = new Set([
  "http://schemas.microsoft.com/office/drawing/2008/diagram",
]);
const DIAGRAM_DATA_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
const DIAGRAM_DRAWING_RELATIONSHIP_TYPE =
  "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
const DOCUMENT_PART_PATH = "word/document.xml";
const WORD_DRAWING_NAMESPACE_URIS = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
]);
const crc32 = (bytes: Uint8Array): number => {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
};

const adler32 = (bytes: Uint8Array): number => {
  let a = 1;
  let b = 0;
  for (const byte of bytes) {
    a = (a + byte) % 65521;
    b = (b + a) % 65521;
  }
  return (b << 16) | a;
};

const pngChunk = (type: string, data: Uint8Array): Uint8Array => {
  const name = new TextEncoder().encode(type);
  const output = new Uint8Array(12 + data.length);
  const view = new DataView(output.buffer);
  view.setUint32(0, data.length);
  output.set(name, 4);
  output.set(data, 8);
  view.setUint32(8 + data.length, crc32(output.subarray(4, 8 + data.length)));
  return output;
};

const previewPng = (width: number, height: number, shapes: PreviewShape[]): Uint8Array => {
  const scale = Math.min(1, Math.sqrt(MAX_PREVIEW_PIXELS / (width * height)));
  const w = Math.max(1, Math.round(width * scale));
  const h = Math.max(1, Math.round(height * scale));
  const pixels = new Uint8Array((w * 4 + 1) * h);
  for (let y = 0; y < h; y += 1) {
    const row = y * (w * 4 + 1);
    pixels[row] = 0;
    for (let x = 0; x < w; x += 1) {
      const index = row + 1 + x * 4;
      pixels[index] = 238;
      pixels[index + 1] = 242;
      pixels[index + 2] = 247;
      pixels[index + 3] = 255;
    }
  }
  let paintedPixels = 0;
  for (const shape of shapes) {
    const sx = Math.max(0, Math.round((shape.x / width) * w));
    const sy = Math.max(0, Math.round((shape.y / height) * h));
    const ex = Math.min(w, Math.round(((shape.x + shape.width) / width) * w));
    const ey = Math.min(h, Math.round(((shape.y + shape.height) / height) * h));
    const red = Number.parseInt(shape.color.slice(0, 2), 16) || 232;
    const green = Number.parseInt(shape.color.slice(2, 4), 16) || 238;
    const blue = Number.parseInt(shape.color.slice(4, 6), 16) || 247;
    const shapePixels = Math.max(0, ex - sx) * Math.max(0, ey - sy);
    if (paintedPixels + shapePixels > MAX_PREVIEW_PAINT_PIXELS) {
      break;
    }
    paintedPixels += shapePixels;
    for (let y = sy; y < ey; y += 1) {
      const row = y * (w * 4 + 1);
      for (let x = sx; x < ex; x += 1) {
        const index = row + 1 + x * 4;
        pixels[index] = red;
        pixels[index + 1] = green;
        pixels[index + 2] = blue;
      }
    }
  }
  const compressed = new Uint8Array(2 + pixels.length + Math.ceil(pixels.length / 65_535) * 5 + 4);
  compressed[0] = 0x78;
  compressed[1] = 0x01;
  let cursor = 2;
  for (let offset = 0; offset < pixels.length;) {
    const length = Math.min(65_535, pixels.length - offset);
    compressed[cursor++] = offset + length === pixels.length ? 1 : 0;
    compressed[cursor++] = length & 255;
    compressed[cursor++] = length >>> 8;
    compressed[cursor++] = ~length & 255;
    compressed[cursor++] = (~length >>> 8) & 255;
    compressed.set(pixels.subarray(offset, offset + length), cursor);
    cursor += length;
    offset += length;
  }
  const output = new Uint8Array(cursor + 4);
  output.set(compressed.subarray(0, cursor));
  const adler = new DataView(output.buffer);
  adler.setUint32(cursor, adler32(pixels));
  const signature = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
  const header = new Uint8Array(13);
  new DataView(header.buffer).setUint32(0, w);
  new DataView(header.buffer).setUint32(4, h);
  header[8] = 8;
  header[9] = 6;
  const chunks = [
    signature,
    pngChunk("IHDR", header),
    pngChunk("IDAT", output),
    pngChunk("IEND", new Uint8Array()),
  ];
  const png = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.length, 0));
  let position = 0;
  for (const chunk of chunks) {
    png.set(chunk, position);
    position += chunk.length;
  }
  return png;
};

const extent = (drawing: XmlElement): { width: number; height: number } => {
  const container =
    findChildByNamespaceUri(drawing, WORD_DRAWING_NAMESPACE_URIS, "inline") ??
    findChildByNamespaceUri(drawing, WORD_DRAWING_NAMESPACE_URIS, "anchor");
  const value = findChildByNamespaceUri(container, WORD_DRAWING_NAMESPACE_URIS, "extent");
  return {
    width: parseNumericAttribute(value, null, "cx") ?? 0,
    height: parseNumericAttribute(value, null, "cy") ?? 0,
  };
};

type PreviewShape = { x: number; y: number; width: number; height: number; color: string };

/** Parse the part a relationship of the given type names, by that relationship's id. */
const partByRelationshipId = ({
  rels,
  media,
  rId,
  type,
}: {
  rels: RelationshipMap;
  media: Map<string, MediaFile>;
  rId: string | null;
  type: string;
}): XmlElement | null => {
  const resolved = resolveRelationshipIdOfType(rels, rId ?? undefined, type);
  if (resolved.status !== "resolved" || !resolved.relationship.target) {
    return null;
  }
  const file = media.get(resolveRelativePath(DOCUMENT_PART_PATH, resolved.relationship.target));
  return file?.data ? parseXmlDocument(new TextDecoder().decode(file.data)) : null;
};

/**
 * The drawing cache this diagram points at, reached through its own ids.
 *
 * `dgm:relIds/@r:dm` names the data part, and that part's `dsp:dataModelExt`
 * extension names the cached drawing. Scanning the relationship map for a type
 * instead of following the ids reads the wrong diagram whenever a document has
 * more than one, which is why the scan refused outright on a second match:
 * both diagrams in a two-diagram document then got no preview at all.
 */
const cachedDiagramShapes = (
  graphicData: XmlElement,
  rels: RelationshipMap,
  media: Map<string, MediaFile>,
): PreviewShape[] => {
  const relIds = findChildByNamespaceUri(graphicData, DIAGRAM_NAMESPACE_URIS, "relIds");
  const data = partByRelationshipId({
    rels,
    media,
    rId: getAttributeByNamespaceUri(relIds, OFFICE_RELATIONSHIP_NAMESPACE_URIS, "dm"),
    type: DIAGRAM_DATA_RELATIONSHIP_TYPE,
  });
  if (!data) {
    return [];
  }
  const dataModelExt = descendantsByNamespace(
    data,
    DIAGRAM_DRAWING_NAMESPACE_URIS,
    "dataModelExt",
  ).at(0);
  const root = partByRelationshipId({
    rels,
    media,
    rId: getAttribute(dataModelExt, null, "relId"),
    type: DIAGRAM_DRAWING_RELATIONSHIP_TYPE,
  });
  if (!root) {
    return [];
  }
  const shapes: PreviewShape[] = [];
  for (const shape of descendantsByNamespace(root, DIAGRAM_DRAWING_NAMESPACE_URIS, "sp").slice(
    0,
    MAX_PREVIEW_SHAPES,
  )) {
    const properties = findChildByNamespaceUri(shape, DIAGRAM_DRAWING_NAMESPACE_URIS, "spPr");
    const transform = findChildByNamespaceUri(properties, DRAWINGML_NAMESPACE_URIS, "xfrm");
    const offset = findChildByNamespaceUri(transform, DRAWINGML_NAMESPACE_URIS, "off");
    const size = findChildByNamespaceUri(transform, DRAWINGML_NAMESPACE_URIS, "ext");
    const width = parseNumericAttribute(size, null, "cx") ?? 0;
    const height = parseNumericAttribute(size, null, "cy") ?? 0;
    if (width <= 0 || height <= 0) {
      continue;
    }
    const fill = findChildByNamespaceUri(properties, DRAWINGML_NAMESPACE_URIS, "solidFill");
    const color = findChildByNamespaceUri(fill, DRAWINGML_NAMESPACE_URIS, "srgbClr");
    shapes.push({
      x: parseNumericAttribute(offset, null, "x") ?? 0,
      y: parseNumericAttribute(offset, null, "y") ?? 0,
      width,
      height,
      color: getAttribute(color, null, "val") ?? "E8EEF7",
    });
  }
  return shapes;
};

const descendantsByNamespace = (
  root: XmlElement,
  namespaceUris: ReadonlySet<string>,
  localName: string,
): XmlElement[] => {
  const result: XmlElement[] = [];
  const visit = (element: XmlElement): void => {
    if (
      element.namespaceUri &&
      namespaceUris.has(element.namespaceUri) &&
      getLocalName(element.name ?? "") === localName
    ) {
      result.push(element);
    }
    for (const child of element.elements ?? []) {
      if (child.type === "element") {
        visit(child);
      }
    }
  };
  visit(root);
  return result;
};

/** Create a deliberately simple, bounded preview; it is never an editable diagram projection. */
export const parseDiagramPreview = (
  drawing: XmlElement,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined,
): Image | null => {
  if (!rels || !media) {
    return null;
  }
  const graphicData = descendantsByNamespace(drawing, DRAWINGML_NAMESPACE_URIS, "graphicData").at(
    0,
  );
  if (!graphicData || !DIAGRAM_NAMESPACE_URIS.has(getAttribute(graphicData, null, "uri") ?? "")) {
    return null;
  }
  const { width, height } = extent(drawing);
  if (width <= 0 || height <= 0) {
    return null;
  }
  const png = previewPng(width, height, cachedDiagramShapes(graphicData, rels, media));
  let binary = "";
  for (let offset = 0; offset < png.length; offset += 0x8000) {
    binary += String.fromCodePoint(...png.subarray(offset, offset + 0x8000));
  }
  const image: Image = {
    type: "image",
    rId: "",
    src: `data:image/png;base64,${btoa(binary)}`,
    mimeType: "image/png",
    filename: "smartart-preview.png",
    size: { width, height },
    wrap: { type: "inline" },
  };
  return image;
};
