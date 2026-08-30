import {
  findChild,
  findDeep,
  getAttribute,
  getChildElements,
  getLocalName,
  type XmlElement,
} from "./xmlParser";

const MAX_VML_PREVIEW_DEPTH = 16;
const MAX_VML_PREVIEW_ELEMENTS = 256;
const MAX_VML_PREVIEW_PATH_POINTS = 20_000;
const MAX_VML_PREVIEW_COORDINATE = 1_000_000;
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
const PROTOTYPE_POLLUTION_KEYS = new Set(["__proto__", "constructor", "prototype"]);

export type VmlPreviewResult =
  | {
      type: "rendered";
      svg: string;
      widthPx: number;
      heightPx: number;
      style: Record<string, string>;
    }
  | { type: "empty" }
  | { type: "invalid" };

type PreviewFragmentResult =
  | { type: "painted"; svg: string }
  | { type: "empty" }
  | { type: "invalid" };

type PreviewBudget = {
  elements: number;
  pathPoints: number;
  svgCharacters: number;
};

type CoordinatePair = { x: number; y: number };

type LocalCoordinateSpace = {
  origin: CoordinatePair;
  size: CoordinatePair;
};

type CoordinateViewport = CoordinatePair & {
  width: number;
  height: number;
};

type PathState =
  | { type: "idle" }
  | {
      type: "positioned";
      x: number;
      y: number;
      commands: string[];
    }
  | { type: "drawing"; x: number; y: number; commands: string[] }
  | { type: "ended-empty" }
  | { type: "ended-painted"; commands: string[] };

type PathResult = { type: "painted"; pathData: string } | { type: "empty" } | { type: "invalid" };

type ShapePaint = { type: "painted"; fill: string; stroke: string } | { type: "empty" };

const EMPTY_RESULT = { type: "empty" } as const;
const INVALID_RESULT = { type: "invalid" } as const;

type VmlPathCommand = "e" | "l" | "m" | "r";

const isVmlPathCommand = (value: string): value is VmlPathCommand => {
  switch (value) {
    case "e":
    case "l":
    case "m":
    case "r":
      return true;
    default:
      return false;
  }
};

/** Parse a VML/CSS length into CSS pixels without accepting partial numbers. */
export const vmlCssLengthToPx = (raw: string | undefined): number | undefined => {
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
  if (!Number.isFinite(value)) {
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
};

/** Read a VML `style` attribute into a prototype-free lowercased record. */
export const parseVmlStyle = (style: string | null): Record<string, string> => {
  const result: Record<string, string> = Object.create(null);
  if (!style) {
    return result;
  }
  for (const declaration of style.split(";")) {
    const colon = declaration.indexOf(":");
    if (colon < 0) {
      continue;
    }
    const key = declaration.slice(0, colon).trim().toLowerCase();
    const value = declaration.slice(colon + 1).trim();
    if (!key || PROTOTYPE_POLLUTION_KEYS.has(key)) {
      continue;
    }
    result[key] = value;
  }
  return result;
};

/** Parse a finite VML number without exponent or partial-number coercion. */
export const parseVmlNumber = (raw: string | null | undefined): number | undefined => {
  if (!raw || !/^-?(?:\d+(?:\.\d+)?|\.\d+)$/u.test(raw.trim())) {
    return undefined;
  }
  const value = Number.parseFloat(raw);
  return Number.isFinite(value) ? value : undefined;
};

/** Check a physical preview dimension against the retained SVG layout bound. */
export const isValidVmlPreviewDimension = (value: number | undefined): value is number =>
  value !== undefined && value > 0 && value <= MAX_VML_PREVIEW_DIMENSION_PX;

/** Encode a generated, already-sanitized SVG without exceeding its output cap. */
export const vmlSvgDataUrl = (svg: string): string | undefined =>
  svg.length <= MAX_VML_SVG_CHARACTERS
    ? `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`
    : undefined;

const validCoordinate = (value: number | undefined): value is number =>
  value !== undefined && Math.abs(value) <= MAX_VML_PREVIEW_COORDINATE;

const validExtent = (value: number | undefined): value is number =>
  value !== undefined && value > 0 && value <= MAX_VML_PREVIEW_COORDINATE;

const coordinatePair = (raw: string | null): CoordinatePair | undefined => {
  const [xRaw, yRaw, extra] = raw?.split(",").map((part) => part.trim()) ?? [];
  if (extra !== undefined) {
    return undefined;
  }
  const x = parseVmlNumber(xRaw);
  const y = parseVmlNumber(yRaw);
  return validCoordinate(x) && validCoordinate(y) ? { x, y } : undefined;
};

const localCoordinateSpace = (
  element: XmlElement,
  fallbackWidth: number,
  fallbackHeight: number,
): LocalCoordinateSpace | undefined => {
  const originRaw = getAttribute(element, null, "coordorigin");
  const sizeRaw = getAttribute(element, null, "coordsize");
  const origin = originRaw ? coordinatePair(originRaw) : { x: 0, y: 0 };
  const size = sizeRaw ? coordinatePair(sizeRaw) : { x: fallbackWidth, y: fallbackHeight };
  if (!origin || !size) {
    return undefined;
  }
  return validExtent(size.x) && validExtent(size.y) ? { origin, size } : undefined;
};

const coordinateViewport = (style: Record<string, string>): CoordinateViewport | undefined => {
  const x = parseVmlNumber(style["left"]) ?? 0;
  const y = parseVmlNumber(style["top"]) ?? 0;
  const width = parseVmlNumber(style["width"]);
  const height = parseVmlNumber(style["height"]);
  return validCoordinate(x) && validCoordinate(y) && validExtent(width) && validExtent(height)
    ? { x, y, width, height }
    : undefined;
};

const transformMatrix = (
  viewport: CoordinateViewport,
  space: LocalCoordinateSpace,
): string | undefined => {
  const scaleX = viewport.width / space.size.x;
  const scaleY = viewport.height / space.size.y;
  const translateX = viewport.x - space.origin.x * scaleX;
  const translateY = viewport.y - space.origin.y * scaleY;
  const values = [scaleX, scaleY, translateX, translateY];
  if (
    values.some((value) => !Number.isFinite(value) || Math.abs(value) > MAX_VML_PREVIEW_COORDINATE)
  ) {
    return undefined;
  }
  return `matrix(${scaleX} 0 0 ${scaleY} ${translateX} ${translateY})`;
};

const safeColor = (raw: string | null, fallback: string): string => {
  const normalized = raw?.trim().toLowerCase();
  if (!normalized) {
    return fallback;
  }
  if (normalized === "none") {
    return "none";
  }
  if (SAFE_VML_COLORS.has(normalized)) {
    return normalized;
  }
  if (/^#[0-9a-f]{3}(?:[0-9a-f]{3})?$/u.test(normalized)) {
    return normalized;
  }
  return /^[0-9a-f]{6}$/u.test(normalized) ? `#${normalized}` : fallback;
};

const isVmlFalse = (raw: string | null): boolean => {
  switch (raw?.trim().toLowerCase()) {
    case "0":
    case "f":
    case "false":
      return true;
    default:
      return false;
  }
};

const shapePaint = (shape: XmlElement): ShapePaint => {
  const fill = isVmlFalse(getAttribute(shape, null, "filled"))
    ? "none"
    : safeColor(getAttribute(shape, null, "fillcolor"), "white");
  const stroke = isVmlFalse(getAttribute(shape, null, "stroked"))
    ? "none"
    : safeColor(getAttribute(shape, null, "strokecolor"), "black");
  return fill === "none" && stroke === "none" ? EMPTY_RESULT : { type: "painted", fill, stroke };
};

const isHidden = (style: Record<string, string>): boolean =>
  style["visibility"]?.toLowerCase() === "hidden" || style["display"]?.toLowerCase() === "none";

const consumeSvgCharacters = (budget: PreviewBudget, count: number): boolean => {
  if (budget.svgCharacters + count > MAX_VML_SVG_CHARACTERS) {
    return false;
  }
  budget.svgCharacters += count;
  return true;
};

const consumePathPoint = (budget: PreviewBudget): boolean => {
  if (budget.pathPoints >= MAX_VML_PREVIEW_PATH_POINTS) {
    return false;
  }
  budget.pathPoints += 1;
  return true;
};

const pathArgumentPairs = (raw: string): CoordinatePair[] | undefined => {
  const parts = raw.split(",").map((part) => part.trim());
  if (parts.length === 0 || parts.length % 2 !== 0) {
    return undefined;
  }
  const pairs: CoordinatePair[] = [];
  for (let index = 0; index < parts.length; index += 2) {
    const xRaw = parts.at(index);
    const yRaw = parts.at(index + 1);
    const x = xRaw === "" ? 0 : parseVmlNumber(xRaw);
    const y = yRaw === "" ? 0 : parseVmlNumber(yRaw);
    if (!validCoordinate(x) || !validCoordinate(y)) {
      return undefined;
    }
    pairs.push({ x, y });
  }
  return pairs;
};

const appendPathCommand = (commands: string[], command: string, budget: PreviewBudget): boolean => {
  if (!consumeSvgCharacters(budget, command.length + 1)) {
    return false;
  }
  commands.push(command);
  return true;
};

const renderVmlPath = (raw: string, budget: PreviewBudget): PathResult => {
  if (raw.length > MAX_VML_SVG_CHARACTERS) {
    return INVALID_RESULT;
  }
  let state: PathState = { type: "idle" };
  let offset = 0;
  while (offset < raw.length) {
    while (/\s/u.test(raw.charAt(offset))) {
      offset += 1;
    }
    if (offset >= raw.length) {
      break;
    }
    const command = raw.charAt(offset).toLowerCase();
    if (!isVmlPathCommand(command)) {
      return INVALID_RESULT;
    }
    offset += 1;
    const argumentStart = offset;
    while (offset < raw.length && !/[a-z]/iu.test(raw.charAt(offset))) {
      offset += 1;
    }
    const rawArguments = raw.slice(argumentStart, offset).trim();
    if (state.type === "ended-empty" || state.type === "ended-painted") {
      return INVALID_RESULT;
    }

    switch (command) {
      case "m": {
        const pairs = pathArgumentPairs(rawArguments);
        const point = pairs?.at(0);
        if (state.type !== "idle" || pairs?.length !== 1 || !point || !consumePathPoint(budget)) {
          return INVALID_RESULT;
        }
        const commands: string[] = [];
        if (!appendPathCommand(commands, `M ${point.x} ${point.y}`, budget)) {
          return INVALID_RESULT;
        }
        state = {
          type: "positioned",
          x: point.x,
          y: point.y,
          commands,
        };
        break;
      }
      case "l":
      case "r": {
        if (state.type !== "positioned" && state.type !== "drawing") {
          return INVALID_RESULT;
        }
        const pairs = pathArgumentPairs(rawArguments);
        if (!pairs) {
          return INVALID_RESULT;
        }
        let x: number = state.x;
        let y: number = state.y;
        for (const point of pairs) {
          x = command === "r" ? x + point.x : point.x;
          y = command === "r" ? y + point.y : point.y;
          if (!validCoordinate(x) || !validCoordinate(y) || !consumePathPoint(budget)) {
            return INVALID_RESULT;
          }
          if (!appendPathCommand(state.commands, `L ${x} ${y}`, budget)) {
            return INVALID_RESULT;
          }
        }
        state = {
          type: "drawing",
          x,
          y,
          commands: state.commands,
        };
        break;
      }
      case "e": {
        if (rawArguments || state.type === "idle") {
          return INVALID_RESULT;
        }
        state =
          state.type === "positioned"
            ? { type: "ended-empty" }
            : { type: "ended-painted", commands: state.commands };
        break;
      }
      default:
        command satisfies never;
    }
  }

  switch (state.type) {
    case "ended-painted":
      return { type: "painted", pathData: state.commands.join(" ") };
    case "ended-empty":
      return EMPTY_RESULT;
    case "idle":
    case "positioned":
    case "drawing":
      return INVALID_RESULT;
    default:
      return state satisfies never;
  }
};

const primitiveShapeFragment = (
  shape: XmlElement,
  viewport: CoordinateViewport,
  budget: PreviewBudget,
): PreviewFragmentResult => {
  const paint = shapePaint(shape);
  if (paint.type === "empty") {
    return EMPTY_RESULT;
  }
  const localName = getLocalName(shape.name ?? "");
  let svg: string;
  if (localName === "oval") {
    svg = `<ellipse cx="${viewport.x + viewport.width / 2}" cy="${viewport.y + viewport.height / 2}" rx="${viewport.width / 2}" ry="${viewport.height / 2}" fill="${paint.fill}" stroke="${paint.stroke}"/>`;
  } else {
    const radius = localName === "roundrect" ? Math.min(viewport.width, viewport.height) * 0.1 : 0;
    svg = `<rect x="${viewport.x}" y="${viewport.y}" width="${viewport.width}" height="${viewport.height}" rx="${radius}" fill="${paint.fill}" stroke="${paint.stroke}"/>`;
  }
  return consumeSvgCharacters(budget, svg.length) ? { type: "painted", svg } : INVALID_RESULT;
};

const lineFragment = (line: XmlElement, budget: PreviewBudget): PreviewFragmentResult => {
  if (isVmlFalse(getAttribute(line, null, "stroked"))) {
    return EMPTY_RESULT;
  }
  const from = coordinatePair(getAttribute(line, null, "from"));
  const to = coordinatePair(getAttribute(line, null, "to"));
  if (!from || !to) {
    return INVALID_RESULT;
  }
  const stroke = safeColor(getAttribute(line, null, "strokecolor"), "black");
  const svg = `<line x1="${from.x}" y1="${from.y}" x2="${to.x}" y2="${to.y}" stroke="${stroke}"/>`;
  return consumeSvgCharacters(budget, svg.length) ? { type: "painted", svg } : INVALID_RESULT;
};

const pathShapeFragment = (shape: XmlElement, budget: PreviewBudget): PreviewFragmentResult => {
  if (findChild(shape, "v", "textbox") || findChild(shape, "v", "imagedata")) {
    return EMPTY_RESULT;
  }
  const style = parseVmlStyle(getAttribute(shape, null, "style"));
  if (isHidden(style)) {
    return EMPTY_RESULT;
  }
  const paint = shapePaint(shape);
  if (paint.type === "empty") {
    return EMPTY_RESULT;
  }
  const rawPath = getAttribute(shape, null, "path");
  if (!rawPath) {
    return EMPTY_RESULT;
  }
  const viewport = coordinateViewport(style);
  if (!viewport) {
    return INVALID_RESULT;
  }
  const space = localCoordinateSpace(shape, viewport.width, viewport.height);
  const transform = space ? transformMatrix(viewport, space) : undefined;
  if (!transform) {
    return INVALID_RESULT;
  }
  const path = renderVmlPath(rawPath, budget);
  if (path.type !== "painted") {
    return path;
  }
  const svg = `<path transform="${transform}" d="${path.pathData}" fill="${paint.fill}" stroke="${paint.stroke}"/>`;
  const wrapperCharacters = svg.length - path.pathData.length;
  return consumeSvgCharacters(budget, wrapperCharacters)
    ? { type: "painted", svg }
    : INVALID_RESULT;
};

const renderGroupChildren = (
  group: XmlElement,
  budget: PreviewBudget,
  depth: number,
): PreviewFragmentResult => {
  const fragments: string[] = [];
  for (const child of getChildElements(group)) {
    if (budget.elements >= MAX_VML_PREVIEW_ELEMENTS) {
      return INVALID_RESULT;
    }
    budget.elements += 1;
    const localName = getLocalName(child.name ?? "");
    let result: PreviewFragmentResult;
    switch (localName) {
      case "group": {
        if (depth >= MAX_VML_PREVIEW_DEPTH) {
          return INVALID_RESULT;
        }
        const style = parseVmlStyle(getAttribute(child, null, "style"));
        if (isHidden(style)) {
          result = EMPTY_RESULT;
          break;
        }
        const viewport = coordinateViewport(style);
        const space = viewport
          ? localCoordinateSpace(child, viewport.width, viewport.height)
          : undefined;
        const transform = viewport && space ? transformMatrix(viewport, space) : undefined;
        if (!transform) {
          return INVALID_RESULT;
        }
        const content = renderGroupChildren(child, budget, depth + 1);
        if (content.type !== "painted") {
          result = content;
          break;
        }
        const svg = `<g transform="${transform}">${content.svg}</g>`;
        const wrapperCharacters = svg.length - content.svg.length;
        result = consumeSvgCharacters(budget, wrapperCharacters)
          ? { type: "painted", svg }
          : INVALID_RESULT;
        break;
      }
      case "shape":
        result = pathShapeFragment(child, budget);
        break;
      case "line":
        result = lineFragment(child, budget);
        break;
      case "oval":
      case "rect":
      case "roundrect": {
        const style = parseVmlStyle(getAttribute(child, null, "style"));
        if (isHidden(style)) {
          result = EMPTY_RESULT;
          break;
        }
        const viewport = coordinateViewport(style);
        result = viewport ? primitiveShapeFragment(child, viewport, budget) : INVALID_RESULT;
        break;
      }
      default:
        result = EMPTY_RESULT;
    }
    switch (result.type) {
      case "painted":
        fragments.push(result.svg);
        break;
      case "empty":
        break;
      case "invalid":
        return INVALID_RESULT;
      default:
        result satisfies never;
    }
  }
  return fragments.length > 0 ? { type: "painted", svg: fragments.join("") } : EMPTY_RESULT;
};

/** Render a bounded standalone VML rectangle or oval preview. */
export const renderStandaloneVmlPreview = (shape: XmlElement): VmlPreviewResult => {
  if (findDeep(shape, "v", "textbox")) {
    return EMPTY_RESULT;
  }
  const style = parseVmlStyle(getAttribute(shape, null, "style"));
  const widthPx = vmlCssLengthToPx(style["width"]);
  const heightPx = vmlCssLengthToPx(style["height"]);
  if (!isValidVmlPreviewDimension(widthPx) || !isValidVmlPreviewDimension(heightPx)) {
    return INVALID_RESULT;
  }
  const budget: PreviewBudget = { elements: 1, pathPoints: 0, svgCharacters: 0 };
  const content = primitiveShapeFragment(
    shape,
    { x: 0, y: 0, width: widthPx, height: heightPx },
    budget,
  );
  if (content.type !== "painted") {
    return content;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${widthPx} ${heightPx}" width="${widthPx}" height="${heightPx}">${content.svg}</svg>`;
  if (!consumeSvgCharacters(budget, svg.length - content.svg.length)) {
    return INVALID_RESULT;
  }
  return { type: "rendered", svg, widthPx, heightPx, style };
};

/** Render a VML group through bounded, non-clipping local-coordinate transforms. */
export const renderVmlGroupPreview = (group: XmlElement): VmlPreviewResult => {
  const style = parseVmlStyle(getAttribute(group, null, "style"));
  const widthPx = vmlCssLengthToPx(style["width"]);
  const heightPx = vmlCssLengthToPx(style["height"]);
  if (!isValidVmlPreviewDimension(widthPx) || !isValidVmlPreviewDimension(heightPx)) {
    return INVALID_RESULT;
  }
  const space = localCoordinateSpace(group, widthPx, heightPx);
  if (!space) {
    return INVALID_RESULT;
  }
  const budget: PreviewBudget = { elements: 1, pathPoints: 0, svgCharacters: 0 };
  const content = renderGroupChildren(group, budget, 0);
  if (content.type !== "painted") {
    return content;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${space.origin.x} ${space.origin.y} ${space.size.x} ${space.size.y}" width="${widthPx}" height="${heightPx}">${content.svg}</svg>`;
  if (!consumeSvgCharacters(budget, svg.length - content.svg.length)) {
    return INVALID_RESULT;
  }
  return { type: "rendered", svg, widthPx, heightPx, style };
};
