import type { Image, MediaFile, PreviewShape, RelationshipMap } from "../types/document";
import { PREVIEW_KINDS } from "./previewBudget";
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

/**
 * How many `dsp:sp` a preview describes at most.
 *
 * The bound is on the description rather than on any picture built from it:
 * the shapes are what a backend draws, and a drawing with ten thousand of them
 * is a drawing folio summarises rather than reproduces.
 */
export const MAX_PREVIEW_SHAPES = 128;

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

/**
 * Diagram parts already read during this parse, keyed by package path.
 *
 * Two drawings can name the same data part, and a data part can name a drawing
 * part another data part also names. Without this the same XML is decoded and
 * parsed once per referring drawing, which is quadratic in the packages that
 * reuse a cache. The map lives for one call into {@link parseDiagramPreviews}
 * and is bounded by the number of parts the package holds, so nothing is
 * retained past the parse.
 */
type DiagramPartCache = Map<string, XmlElement | null>;

/** Parse the part a relationship of the given type names, by that relationship's id. */
const partByRelationshipId = ({
  rels,
  media,
  rId,
  type,
  cache,
}: {
  rels: RelationshipMap;
  media: Map<string, MediaFile>;
  rId: string | null;
  type: string;
  cache: DiagramPartCache;
}): XmlElement | null => {
  const resolved = resolveRelationshipIdOfType(rels, rId ?? undefined, type);
  if (resolved.status !== "resolved" || !resolved.relationship.target) {
    return null;
  }
  const path = resolveRelativePath(DOCUMENT_PART_PATH, resolved.relationship.target);
  const cached = cache.get(path);
  if (cached !== undefined) {
    return cached;
  }
  const file = media.get(path);
  const parsed = file?.data ? parseXmlDocument(new TextDecoder().decode(file.data)) : null;
  cache.set(path, parsed);
  return parsed;
};

const matchesNamespace = (
  element: XmlElement,
  namespaceUris: ReadonlySet<string>,
  localName: string,
): boolean =>
  element.namespaceUri !== undefined &&
  namespaceUris.has(element.namespaceUri) &&
  getLocalName(element.name ?? "") === localName;

/**
 * Descendants in document order, stopping once `limit` of them are found.
 *
 * The walk used to collect every match and slice afterwards, so a drawing with
 * ten thousand shapes built a ten-thousand-element array to keep a hundred and
 * twenty-eight of them. The bound belongs in the walk, not after it.
 */
const descendantsByNamespace = (
  root: XmlElement,
  namespaceUris: ReadonlySet<string>,
  localName: string,
  limit: number,
): XmlElement[] => {
  const result: XmlElement[] = [];
  const visit = (element: XmlElement): boolean => {
    if (matchesNamespace(element, namespaceUris, localName)) {
      result.push(element);
      return result.length < limit;
    }
    for (const child of element.elements ?? []) {
      if (child.type === "element" && !visit(child)) {
        return false;
      }
    }
    return true;
  };
  visit(root);
  return result;
};

/** The first match in document order, without walking the rest of the subtree. */
const firstDescendantByNamespace = (
  root: XmlElement,
  namespaceUris: ReadonlySet<string>,
  localName: string,
): XmlElement | null => {
  if (matchesNamespace(root, namespaceUris, localName)) {
    return root;
  }
  for (const child of root.elements ?? []) {
    if (child.type !== "element") {
      continue;
    }
    const found = firstDescendantByNamespace(child, namespaceUris, localName);
    if (found) {
      return found;
    }
  }
  return null;
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
  cache: DiagramPartCache,
): PreviewShape[] => {
  const relIds = findChildByNamespaceUri(graphicData, DIAGRAM_NAMESPACE_URIS, "relIds");
  const data = partByRelationshipId({
    rels,
    media,
    cache,
    rId: getAttributeByNamespaceUri(relIds, OFFICE_RELATIONSHIP_NAMESPACE_URIS, "dm"),
    type: DIAGRAM_DATA_RELATIONSHIP_TYPE,
  });
  if (!data) {
    return [];
  }
  const dataModelExt = firstDescendantByNamespace(
    data,
    DIAGRAM_DRAWING_NAMESPACE_URIS,
    "dataModelExt",
  );
  const root = partByRelationshipId({
    rels,
    media,
    cache,
    rId: getAttribute(dataModelExt, null, "relId"),
    type: DIAGRAM_DRAWING_RELATIONSHIP_TYPE,
  });
  if (!root) {
    return [];
  }
  const shapes: PreviewShape[] = [];
  for (const shape of descendantsByNamespace(
    root,
    DIAGRAM_DRAWING_NAMESPACE_URIS,
    "sp",
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

/**
 * The part cache, keyed by the package's own media map.
 *
 * The cache has to be per package: a global one would answer one document's
 * `word/diagrams/data1.xml` with another's. Keying it on the media map gets
 * that lifetime for free and without threading a reader through every drawing
 * call site, and a weak key means the cache dies with the package rather than
 * retaining its parts.
 */
const PART_CACHES = new WeakMap<Map<string, MediaFile>, DiagramPartCache>();

const partCacheFor = (media: Map<string, MediaFile>): DiagramPartCache => {
  const existing = PART_CACHES.get(media);
  if (existing) {
    return existing;
  }
  const created: DiagramPartCache = new Map();
  PART_CACHES.set(media, created);
  return created;
};

/**
 * Describe a diagram drawing, without drawing it.
 *
 * Deliberately simple and bounded; it is never an editable diagram projection.
 */
export const parseDiagramPreview = (
  drawing: XmlElement,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined,
): Image | null => {
  if (!rels || !media) {
    return null;
  }
  const graphicData = firstDescendantByNamespace(drawing, DRAWINGML_NAMESPACE_URIS, "graphicData");
  if (!graphicData || !DIAGRAM_NAMESPACE_URIS.has(getAttribute(graphicData, null, "uri") ?? "")) {
    return null;
  }
  const { width, height } = extent(drawing);
  if (width <= 0 || height <= 0) {
    return null;
  }
  // No `src`: there is no picture, here or anywhere later. The descriptor says
  // what the drawing looks like, and a backend draws it.
  return {
    type: "image",
    rId: "",
    preview: {
      kind: "diagram",
      extent: { width, height },
      shapes: cachedDiagramShapes(graphicData, rels, media, partCacheFor(media)),
    },
    mimeType: PREVIEW_KINDS.smartArt.mimeType,
    filename: PREVIEW_KINDS.smartArt.filename,
    size: { width, height },
    wrap: { type: "inline" },
  };
};
