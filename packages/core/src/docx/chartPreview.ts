import { decodeText } from "../utils/textEncoding";
import type { Image, MediaFile, PreviewShape, RelationshipMap } from "../types/document";
import { MAX_PREVIEW_POINTS, MAX_PREVIEW_SHAPES } from "./diagramPreview";
import { RELATIONSHIP_TYPES, resolveRelativePath } from "./relsParser";
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

type CachedSeries = { values: (number | undefined)[]; color: string };

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
): (string | undefined)[] | null => {
  const cache = reference ? firstDescendant(reference, CHART, cacheName) : null;
  if (!cache) return null;
  const declaredCount = parseNumericAttribute(child(cache, "ptCount"), null, "val");
  if (declaredCount === undefined || !Number.isInteger(declaredCount) || declaredCount <= 0)
    return null;
  const values = Array<string | undefined>(Math.min(declaredCount, MAX_POINTS));
  const seen = new Set<number>();
  for (const point of getChildElements(cache).filter((item) => matches(item, CHART, "pt"))) {
    const index = parseNumericAttribute(point, null, "idx");
    if (
      index === undefined ||
      !Number.isInteger(index) ||
      index < 0 ||
      index >= declaredCount ||
      seen.has(index)
    )
      return null;
    seen.add(index);
    const value = child(point, "v");
    const text = value?.elements?.find((item) => item.type === "text")?.text;
    if (text === undefined) return null;
    if (index < values.length) values[index] = String(text);
  }
  return values;
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
    const values = rawValues.map((value) => (value === undefined ? undefined : Number(value)));
    if (values.some((value) => value !== undefined && !Number.isFinite(value))) return null;
    result.push({ values, color: seriesColor(series, index) });
  }
  return result;
};

const pushShape = (shapes: PreviewShape[], shape: PreviewShape): void => {
  if (shapes.length >= MAX_PREVIEW_SHAPES) return;
  if (
    shape.type !== "polygon" &&
    shape.type !== "polyline" &&
    (shape.x < 0 || shape.y < 0 || shape.width <= 0 || shape.height <= 0)
  )
    return;
  if ((shape.type === "polygon" || shape.type === "polyline") && shape.points.length < 2) return;
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
  const magnitude = Math.max(
    ...series.flatMap((item) => item.values.map((value) => Math.abs(value ?? 0))),
  );
  if (!(magnitude > 0)) return [];
  const normalized = series.map((item) => item.values.map((value) => value && value / magnitude));
  const positives = Array.from({ length: points }, (_, point) =>
    stacked
      ? normalized.reduce((sum, values) => sum + Math.max(0, values[point] ?? 0), 0)
      : Math.max(...normalized.map((values) => Math.max(0, values[point] ?? 0))),
  );
  const negatives = Array.from({ length: points }, (_, point) =>
    stacked
      ? normalized.reduce((sum, values) => sum + Math.min(0, values[point] ?? 0), 0)
      : Math.min(...normalized.map((values) => Math.min(0, values[point] ?? 0))),
  );
  const maximum = Math.max(0, ...positives);
  const minimum = Math.min(0, ...negatives);
  const range = maximum - minimum;
  if (!(range > 0) || points === 0) return [];
  for (let point = 0; point < points; point++) {
    let positive = 0;
    let negative = 0;
    for (const [seriesIndex, item] of series.entries()) {
      const value = normalized[seriesIndex]?.[point];
      if (value === undefined) continue;
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
          type: "rect",
          x: plot.x + Math.min(from, to) * plot.width,
          y: plot.y + offset,
          width: Math.abs(to - from) * plot.width,
          height: cross,
          color: item.color,
        });
      } else {
        const top = 1 - Math.max(from, to);
        pushShape(shapes, {
          type: "rect",
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

const lineShapes = (
  chart: XmlElement,
  series: CachedSeries[],
  width: number,
  height: number,
): PreviewShape[] => {
  const values = series.flatMap((item) => item.values).filter((value) => value !== undefined);
  const magnitude = Math.max(...values.map(Math.abs));
  if (!(magnitude > 0)) return [];
  const normalizedSeries = series.map((item) =>
    item.values.map((value) => value && value / magnitude),
  );
  const normalizedValues = normalizedSeries
    .flatMap((seriesValues) => seriesValues)
    .filter((value) => value !== undefined);
  const minimum = Math.min(...normalizedValues);
  const maximum = Math.max(...normalizedValues);
  const range = maximum - minimum || 1;
  const shapes: PreviewShape[] = [];
  const seriesElements = getChildElements(chart).filter((item) => matches(item, CHART, "ser"));
  for (const [seriesIndex, item] of series.entries()) {
    const denominator = Math.max(1, item.values.length - 1);
    const runs: { x: number; y: number }[][] = [];
    let points: { x: number; y: number }[] = [];
    for (const [index, value] of (normalizedSeries[seriesIndex] ?? []).entries()) {
      if (value === undefined) {
        if (points.length > 0) runs.push(points);
        points = [];
        continue;
      }
      points.push({
        x: width * (0.08 + (0.84 * index) / denominator),
        y: height * (0.9 - (0.8 * (value - minimum)) / range),
      });
    }
    if (points.length > 0) runs.push(points);
    for (const run of runs)
      pushShape(shapes, {
        type: "polyline",
        points: run,
        color: item.color,
        strokeWidth: Math.max(1, Math.min(width, height) * 0.008),
      });
    if (child(seriesElements.at(seriesIndex) ?? null, "marker")) {
      const size = Math.max(1, Math.min(width, height) * 0.018);
      for (const point of runs.flat())
        pushShape(shapes, {
          type: "rect",
          x: point.x - size / 2,
          y: point.y - size / 2,
          width: size,
          height: size,
          color: item.color,
        });
    }
  }
  return shapes;
};

const arcPoints = (
  center: { readonly x: number; readonly y: number },
  radius: number,
  start: number,
  sweep: number,
  segments: number,
) => {
  const points: { x: number; y: number }[] = [];
  for (let step = 0; step <= segments; step += 1) {
    const angle = start + (sweep * step) / segments;
    points.push({ x: center.x + radius * Math.cos(angle), y: center.y + radius * Math.sin(angle) });
  }
  return points;
};

const pieShapes = (
  chart: XmlElement,
  series: CachedSeries[],
  width: number,
  height: number,
  doughnut: boolean,
): PreviewShape[] => {
  const values = series.at(0)?.values ?? [];
  if (values.some((value) => value !== undefined && value < 0)) return [];
  const magnitude = Math.max(...values.map((value) => Math.abs(value ?? 0)));
  if (!(magnitude > 0)) return [];
  const normalized = values.map((value) => value && value / magnitude);
  let total = 0;
  for (const value of normalized) total += value ?? 0;
  if (!(total > 0)) return [];
  const shapes: PreviewShape[] = [];
  const center = { x: width / 2, y: height / 2 };
  const radius = Math.min(width, height) * 0.42;
  const holePercent = parseNumericAttribute(child(chart, "holeSize"), null, "val") ?? 50;
  const hole = doughnut ? radius * Math.min(0.9, Math.max(0.1, holePercent / 100)) : 0;
  let angle =
    ((parseNumericAttribute(child(chart, "firstSliceAng"), null, "val") ?? 0) * Math.PI) / 180 -
    Math.PI / 2;
  for (const [index, value] of normalized.entries()) {
    if (value === undefined || value === 0) continue;
    const sweep = (Math.PI * 2 * value) / total;
    const segments = Math.max(1, Math.min(32, Math.ceil((32 * value) / total)));
    const outer = arcPoints(center, radius, angle, sweep, segments);
    const inner = hole > 0 ? arcPoints(center, hole, angle + sweep, -sweep, segments) : [center];
    pushShape(shapes, {
      type: "polygon",
      points: [...outer, ...inner],
      color: DEFAULT_COLORS[index % DEFAULT_COLORS.length] ?? "4472C4",
    });
    angle += sweep;
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

/** Describe a supported cached chart. AlternateContent fallbacks are resolved by the run parser. */
export const parseChartPreview = (
  drawing: XmlElement,
  rels?: RelationshipMap,
  media?: Map<string, MediaFile>,
): Image | null => {
  if (!rels || !media) return null;
  const part = chartPart(drawing, rels, media);
  const size = extentOf(drawing);
  if (!part || size.width <= 0 || size.height <= 0) return null;
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
    shapes = lineShapes(supported, series, size.width, size.height);
  } else {
    shapes = pieShapes(supported, series, size.width, size.height, kind === "doughnutChart");
  }
  if (shapes.length === 0) return null;
  const pointCount = shapes.reduce(
    (total, shape) =>
      total + (shape.type === "polygon" || shape.type === "polyline" ? shape.points.length : 0),
    0,
  );
  if (pointCount > MAX_PREVIEW_POINTS) return null;
  const finite = shapes.every((shape) =>
    shape.type === "polygon" || shape.type === "polyline"
      ? shape.points.every(({ x, y }) => Number.isFinite(x) && Number.isFinite(y))
      : [shape.x, shape.y, shape.width, shape.height].every(Number.isFinite),
  );
  if (!finite) return null;
  return {
    type: "image",
    preview: { kind: "chart", extent: size, shapes },
    size,
    wrap: { type: "inline" },
  };
};
