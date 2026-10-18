import { decodeText } from "../utils/textEncoding";
import type { Image, MediaFile, PreviewShape, RelationshipMap } from "../types/document";
import { MAX_PREVIEW_SHAPES } from "./diagramPreview";
import { parseRelationships, RELATIONSHIP_TYPES, resolveRelativePath } from "./relsParser";
import {
  findChildByNamespaceUri,
  getAttribute,
  getAttributeByNamespaceUri,
  getChildElements,
  getLocalName,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  parseNumericAttribute,
  parseXmlDocument,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

const DRAWINGML = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/main",
  "http://purl.oclc.org/ooxml/drawingml/main",
]);
const CHART = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/chart",
  "http://purl.oclc.org/ooxml/drawingml/chart",
]);
const WORD_DRAWING = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
]);
const DOCUMENT_PATH = "word/document.xml";
const DEFAULT_COLORS = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"];
const MAX_SERIES = 16;
const MAX_POINTS = 64;

type CachedSeries = { values: number[]; color: string };

const matches = (
  element: XmlElement,
  namespaces: ReadonlySet<string>,
  localName: string,
): boolean =>
  element.namespaceUri !== undefined &&
  namespaces.has(element.namespaceUri) &&
  getLocalName(element.name ?? "") === localName;

const descendants = (
  root: XmlElement,
  namespaces: ReadonlySet<string>,
  localName: string,
): XmlElement[] => {
  const found: XmlElement[] = [];
  const visit = (element: XmlElement): void => {
    if (matches(element, namespaces, localName)) found.push(element);
    for (const child of getChildElements(element)) visit(child);
  };
  visit(root);
  return found;
};

const firstDescendant = (root: XmlElement, namespaces: ReadonlySet<string>, localName: string) =>
  descendants(root, namespaces, localName).at(0) ?? null;

const child = (root: XmlElement | null, localName: string) =>
  findChildByNamespaceUri(root, CHART, localName);

const extentOf = (drawing: XmlElement) => {
  const frame =
    firstDescendant(drawing, WORD_DRAWING, "inline") ??
    firstDescendant(drawing, WORD_DRAWING, "anchor");
  const extent = frame ? findChildByNamespaceUri(frame, WORD_DRAWING, "extent") : null;
  return {
    width: parseNumericAttribute(extent, null, "cx") ?? 0,
    height: parseNumericAttribute(extent, null, "cy") ?? 0,
  };
};

const fileAt = (media: Map<string, MediaFile>, path: string): MediaFile | undefined => {
  const wanted = path.toLowerCase();
  for (const [candidate, file] of media) if (candidate.toLowerCase() === wanted) return file;
  return undefined;
};

const xmlAt = (media: Map<string, MediaFile>, path: string): XmlElement | null => {
  const file = fileAt(media, path);
  return file ? parseXmlDocument(decodeText(new Uint8Array(file.data))) : null;
};

const cacheValues = (
  reference: XmlElement | null,
  cacheName: "numCache" | "strCache",
): string[] | null => {
  const cache = reference ? firstDescendant(reference, CHART, cacheName) : null;
  if (!cache) return null;
  const values: string[] = [];
  for (const point of getChildElements(cache)
    .filter((item) => matches(item, CHART, "pt"))
    .slice(0, MAX_POINTS)) {
    const value = child(point, "v");
    const text = value?.elements?.find((item) => item.type === "text")?.text;
    if (text === undefined) return null;
    values.push(String(text));
  }
  return values.length > 0 ? values : null;
};

const seriesColor = (series: XmlElement, index: number): string => {
  const properties = child(series, "spPr");
  const fill = properties ? firstDescendant(properties, DRAWINGML, "solidFill") : null;
  const color = fill ? firstDescendant(fill, DRAWINGML, "srgbClr") : null;
  const value = getAttribute(color, null, "val");
  return value && /^[0-9a-f]{6}$/iu.test(value)
    ? value
    : (DEFAULT_COLORS[index % DEFAULT_COLORS.length] ?? "4472C4");
};

const cachedSeries = (chart: XmlElement): CachedSeries[] | null => {
  const result: CachedSeries[] = [];
  const seriesElements = getChildElements(chart)
    .filter((item) => matches(item, CHART, "ser"))
    .slice(0, MAX_SERIES);
  if (seriesElements.length === 0) return null;
  for (const [index, series] of seriesElements.entries()) {
    const categoryRef = child(series, "cat");
    const categories = cacheValues(categoryRef, "strCache") ?? cacheValues(categoryRef, "numCache");
    const rawValues = cacheValues(child(series, "val") ?? child(series, "yVal"), "numCache");
    if (!categories || !rawValues || categories.length !== rawValues.length) return null;
    const values = rawValues.map(Number);
    if (values.some((value) => !Number.isFinite(value))) return null;
    result.push({ values, color: seriesColor(series, index) });
  }
  return result;
};

const pushShape = (shapes: PreviewShape[], shape: PreviewShape): void => {
  if (shapes.length >= MAX_PREVIEW_SHAPES) return;
  if (shape.x < 0 || shape.y < 0 || shape.width <= 0 || shape.height <= 0) return;
  shapes.push(shape);
};

const barShapes = (
  chart: XmlElement,
  series: CachedSeries[],
  width: number,
  height: number,
): PreviewShape[] => {
  const shapes: PreviewShape[] = [];
  const plot = { x: width * 0.08, y: height * 0.08, width: width * 0.86, height: height * 0.84 };
  const direction = getAttribute(child(chart, "barDir"), null, "val") ?? "col";
  const stacked =
    (getAttribute(child(chart, "grouping"), null, "val") ?? "clustered") === "stacked";
  const points = series.at(0)?.values.length ?? 0;
  const positives = Array.from({ length: points }, (_, point) =>
    stacked
      ? series.reduce((sum, item) => sum + Math.max(0, item.values[point] ?? 0), 0)
      : Math.max(...series.map((item) => Math.max(0, item.values[point] ?? 0))),
  );
  const negatives = Array.from({ length: points }, (_, point) =>
    stacked
      ? series.reduce((sum, item) => sum + Math.min(0, item.values[point] ?? 0), 0)
      : Math.min(...series.map((item) => Math.min(0, item.values[point] ?? 0))),
  );
  const maximum = Math.max(0, ...positives);
  const minimum = Math.min(0, ...negatives);
  const range = maximum - minimum;
  if (!(range > 0) || points === 0) return [];
  for (let point = 0; point < points; point++) {
    let positive = 0;
    let negative = 0;
    for (const [seriesIndex, item] of series.entries()) {
      const value = item.values[point] ?? 0;
      let start = 0;
      if (stacked) start = value >= 0 ? positive : negative;
      if (value >= 0) positive += value;
      else negative += value;
      const band = (direction === "bar" ? plot.height : plot.width) / points;
      const crossCount = stacked ? 1 : series.length;
      const cross = (band * 0.8) / crossCount;
      const offset = band * point + band * 0.1 + (stacked ? 0 : seriesIndex * cross);
      const from = (start - minimum) / range;
      const to = (start + value - minimum) / range;
      if (value === 0) continue;
      if (direction === "bar") {
        pushShape(shapes, {
          x: plot.x + Math.min(from, to) * plot.width,
          y: plot.y + offset,
          width: Math.abs(to - from) * plot.width,
          height: cross,
          color: item.color,
        });
      } else {
        const top = 1 - Math.max(from, to);
        pushShape(shapes, {
          x: plot.x + offset,
          y: plot.y + top * plot.height,
          width: cross,
          height: Math.abs(to - from) * plot.height,
          color: item.color,
        });
      }
    }
  }
  return shapes;
};

const lineShapes = (series: CachedSeries[], width: number, height: number): PreviewShape[] => {
  const values = series.flatMap((item) => item.values);
  const minimum = Math.min(...values);
  const maximum = Math.max(...values);
  const range = maximum - minimum || 1;
  const shapes: PreviewShape[] = [];
  for (const item of series)
    for (const [index, value] of item.values.entries()) {
      const size = Math.max(1, Math.min(width, height) * 0.018);
      const denominator = Math.max(1, item.values.length - 1);
      pushShape(shapes, {
        x: width * (0.08 + (0.84 * index) / denominator) - size / 2,
        y: height * (0.9 - (0.8 * (value - minimum)) / range) - size / 2,
        width: size,
        height: size,
        color: item.color,
      });
    }
  return shapes;
};

const pieShapes = (series: CachedSeries[], width: number, height: number): PreviewShape[] => {
  const values = series.at(0)?.values ?? [];
  if (values.some((value) => value < 0)) return [];
  const total = values.reduce((sum, value) => sum + value, 0);
  if (!(total > 0)) return [];
  const shapes: PreviewShape[] = [];
  let x = width * 0.08;
  for (const [index, value] of values.entries()) {
    const segmentWidth = (width * 0.84 * value) / total;
    pushShape(shapes, {
      x,
      y: height * 0.18,
      width: segmentWidth,
      height: height * 0.64,
      color: DEFAULT_COLORS[index % DEFAULT_COLORS.length] ?? "4472C4",
    });
    x += segmentWidth;
  }
  return shapes;
};

const chartPart = (drawing: XmlElement, rels: RelationshipMap, media: Map<string, MediaFile>) => {
  const graphicData = firstDescendant(drawing, DRAWINGML, "graphicData");
  if (!graphicData || !CHART.has(getAttribute(graphicData, null, "uri") ?? "")) return null;
  const reference = firstDescendant(graphicData, CHART, "chart");
  const id = getAttributeByNamespaceUri(reference, OFFICE_RELATIONSHIP_NAMESPACE_URIS, "id");
  const relationship = id ? rels.get(id) : undefined;
  if (!relationship || relationship.type !== RELATIONSHIP_TYPES.chart || !relationship.target)
    return null;
  const path = resolveRelativePath(DOCUMENT_PATH, relationship.target);
  const root = xmlAt(media, path);
  return root ? { root, path } : null;
};

/** Resolve the chart part's preview image, used when no user-shape overlay exists. */
const relatedImage = (root: XmlElement, path: string, media: Map<string, MediaFile>) => {
  if (descendants(root, CHART, "userShapes").length > 0) return null;
  const slash = path.lastIndexOf("/");
  const relsPath = `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`;
  const relsFile = fileAt(media, relsPath);
  if (!relsFile) return null;
  const chartRels = parseRelationships(decodeText(new Uint8Array(relsFile.data)));
  const image = [...chartRels.values()].find(
    (relationship) => relationship.type === RELATIONSHIP_TYPES.image,
  );
  if (!image) return null;
  const resolvedPath = resolveRelativePath(path, image.target);
  const file = fileAt(media, resolvedPath);
  const src = file?.dataUrl ?? file?.base64;
  return file && src
    ? { src, mimeType: file.mimeType, ...(file.filename ? { filename: file.filename } : {}) }
    : null;
};

/** Describe a supported cached chart, or its authored related-image fallback. */
export const parseChartPreview = (
  drawing: XmlElement,
  rels?: RelationshipMap,
  media?: Map<string, MediaFile>,
): Image | null => {
  if (!rels || !media) return null;
  const part = chartPart(drawing, rels, media);
  const size = extentOf(drawing);
  if (!part || size.width <= 0 || size.height <= 0) return null;
  const fallback = relatedImage(part.root, part.path, media);
  if (fallback) return { type: "image", ...fallback, size, wrap: { type: "inline" } };
  const plotArea = firstDescendant(part.root, CHART, "plotArea");
  if (!plotArea) return null;
  const supported = getChildElements(plotArea).find(
    (item) =>
      ["barChart", "lineChart", "pieChart", "doughnutChart"].includes(
        getLocalName(item.name ?? ""),
      ) && CHART.has(item.namespaceUri ?? ""),
  );
  if (!supported) return null;
  const series = cachedSeries(supported);
  if (!series) return null;
  const kind = getLocalName(supported.name ?? "");
  let shapes: PreviewShape[];
  if (kind === "barChart") {
    shapes = barShapes(supported, series, size.width, size.height);
  } else if (kind === "lineChart") {
    shapes = lineShapes(series, size.width, size.height);
  } else {
    shapes = pieShapes(series, size.width, size.height);
  }
  if (shapes.length === 0) return null;
  return {
    type: "image",
    preview: { kind: "diagram", extent: size, shapes },
    size,
    wrap: { type: "inline" },
  };
};
