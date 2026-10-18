/**
 * Render a deliberately small, browser-safe WMF subset to generated SVG.
 *
 * Minified bundle cost (measured with `bun build --minify`): 7,002 bytes.
 * Unsupported records fail closed; original WMF bytes remain authoritative.
 */

import {
  assembleMetafileSvg,
  metafileColorRefToCss,
  metafileNumberText,
  transformMetafilePoint,
} from "./metafileSvgShared";

const PLACEABLE_KEY = 0x9ac6_cdd7;
const HEADER_WORDS = 9;
const MAX_WMF_RECORDS = 10_000;
const MAX_WMF_POINTS = 20_000;
const MAX_WMF_DIMENSION = 32_768;
const MAX_WMF_OUTPUT_CHARACTERS = 1_000_000;
const MAX_WMF_STATE_DEPTH = 256;
const MAX_WMF_OBJECTS = 4_096;

const META = {
  eof: 0x0000,
  setBkMode: 0x0102,
  setMapMode: 0x0103,
  setPolyFillMode: 0x0106,
  restoreDc: 0x0127,
  selectObject: 0x012d,
  setWindowOrg: 0x020b,
  setWindowExt: 0x020c,
  setViewportOrg: 0x020d,
  setViewportExt: 0x020e,
  lineTo: 0x0213,
  moveTo: 0x0214,
  createPenIndirect: 0x02fa,
  createBrushIndirect: 0x02fc,
  deleteObject: 0x01f0,
  polygon: 0x0324,
  polyline: 0x0325,
  saveDc: 0x001e,
  ellipse: 0x0418,
  rectangle: 0x041b,
  polyPolygon: 0x0538,
} as const;

const MAP_MODE_TEXT = 1;
const MAP_MODE_ISOTROPIC = 7;
const MAP_MODE_ANISOTROPIC = 8;
const POLY_FILL_ALTERNATE = 1;
const POLY_FILL_WINDING = 2;

type Pen = { type: "pen"; stroke: string | null; width: number };
type Brush = { type: "brush"; fill: string | null };
type GraphicsObject = Pen | Brush;
type Point = { x: number; y: number };
type State = {
  mapMode: typeof MAP_MODE_TEXT | typeof MAP_MODE_ISOTROPIC | typeof MAP_MODE_ANISOTROPIC;
  origin: Point;
  extent: Point | null;
  viewportOrigin: Point;
  viewportExtent: Point | null;
  current: Point;
  fillRule: "evenodd" | "nonzero";
  pen: Pen | null;
  brush: Brush | null;
};

const cloneState = (state: State): State => ({
  ...state,
  origin: { ...state.origin },
  extent: state.extent ? { ...state.extent } : null,
  viewportOrigin: { ...state.viewportOrigin },
  viewportExtent: state.viewportExtent ? { ...state.viewportExtent } : null,
  current: { ...state.current },
});

const word = (view: DataView, offset: number, end: number): number | undefined =>
  offset >= 0 && offset + 2 <= end ? view.getUint16(offset, true) : undefined;
const short = (view: DataView, offset: number, end: number): number | undefined =>
  offset >= 0 && offset + 2 <= end ? view.getInt16(offset, true) : undefined;
const dword = (view: DataView, offset: number, end: number): number | undefined =>
  offset >= 0 && offset + 4 <= end ? view.getUint32(offset, true) : undefined;

const mappedPoint = (state: State, point: Point): Point | undefined => {
  if (state.mapMode === MAP_MODE_TEXT) {
    return {
      x: state.viewportOrigin.x + point.x - state.origin.x,
      y: state.viewportOrigin.y + point.y - state.origin.y,
    };
  }
  const extent = state.extent;
  const viewportExtent = state.viewportExtent;
  if (!extent || !viewportExtent || extent.x === 0 || extent.y === 0) return undefined;
  if (state.mapMode === MAP_MODE_ISOTROPIC) {
    const scale = Math.min(
      Math.abs(viewportExtent.x / extent.x),
      Math.abs(viewportExtent.y / extent.y),
    );
    return {
      x:
        state.viewportOrigin.x +
        (point.x - state.origin.x) * Math.sign(viewportExtent.x / extent.x) * scale,
      y:
        state.viewportOrigin.y +
        (point.y - state.origin.y) * Math.sign(viewportExtent.y / extent.y) * scale,
    };
  }
  return transformMetafilePoint({
    point,
    windowOrigin: state.origin,
    windowExtent: extent,
    viewportOrigin: state.viewportOrigin,
    viewportExtent,
    outputOrigin: { x: 0, y: 0 },
    maxCoordinate: MAX_WMF_DIMENSION * 100,
  });
};

const mappedPenWidth = (state: State, pen: Pen | null): number | undefined => {
  const width = pen?.width ?? 0;
  if (width === 0) return 1;
  if (state.mapMode === MAP_MODE_TEXT) return width;
  const extent = state.extent;
  const viewportExtent = state.viewportExtent;
  if (!extent || !viewportExtent || extent.x === 0 || extent.y === 0) return undefined;
  const xScale = Math.abs(viewportExtent.x / extent.x);
  const yScale = Math.abs(viewportExtent.y / extent.y);
  const scale =
    state.mapMode === MAP_MODE_ISOTROPIC ? Math.min(xScale, yScale) : (xScale + yScale) / 2;
  // WMF geometric pens have one logical width; average the device-axis scales for anisotropic maps.
  const mapped = width * scale;
  return Number.isFinite(mapped) ? mapped : undefined;
};

const pathPoint = (command: string, point: Point): string =>
  `${command}${metafileNumberText(point.x)} ${metafileNumberText(point.y)}`;

const paint = (state: State): string | undefined => {
  if (!state.brush && !state.pen) return undefined;
  const fill = state.brush?.fill ?? "none";
  const stroke = state.pen?.stroke ?? "none";
  const width = mappedPenWidth(state, state.pen);
  if (width === undefined) return undefined;
  return `fill="${fill}" stroke="${stroke}" stroke-width="${metafileNumberText(width)}"`;
};

const readPoint = (view: DataView, offset: number, end: number): Point | undefined => {
  const x = short(view, offset, end);
  const y = short(view, offset + 2, end);
  return x === undefined || y === undefined ? undefined : { x, y };
};

/** Convert the supported WMF vector subset to inert generated SVG. */
export function renderWmfSvg(data: ArrayBuffer | Uint8Array): string | null {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (bytes.byteLength < 18) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let headerOffset = 0;
  let placeableOrigin: Point | null = null;
  let placeableExtent: Point | null = null;
  if (dword(view, 0, bytes.byteLength) === PLACEABLE_KEY) {
    if (bytes.byteLength < 40) return null;
    const left = short(view, 6, 22);
    const top = short(view, 8, 22);
    const right = short(view, 10, 22);
    const bottom = short(view, 12, 22);
    let checksum = 0;
    for (let offset = 0; offset < 20; offset += 2) checksum ^= view.getUint16(offset, true);
    if (
      left === undefined ||
      top === undefined ||
      right === undefined ||
      bottom === undefined ||
      checksum !== word(view, 20, 22)
    )
      return null;
    placeableOrigin = { x: left, y: top };
    placeableExtent = { x: right - left, y: bottom - top };
    headerOffset = 22;
  }

  const type = word(view, headerOffset, bytes.byteLength);
  const headerWords = word(view, headerOffset + 2, bytes.byteLength);
  const version = word(view, headerOffset + 4, bytes.byteLength);
  const fileWords = dword(view, headerOffset + 6, bytes.byteLength);
  const objectCount = word(view, headerOffset + 10, bytes.byteLength);
  const maxRecordWords = dword(view, headerOffset + 12, bytes.byteLength);
  const parameterCount = word(view, headerOffset + 16, bytes.byteLength);
  if (
    (type !== 1 && type !== 2) ||
    headerWords !== HEADER_WORDS ||
    (version !== 0x0100 && version !== 0x0300) ||
    fileWords === undefined ||
    fileWords * 2 !== bytes.byteLength - headerOffset ||
    objectCount === undefined ||
    objectCount > MAX_WMF_OBJECTS ||
    maxRecordWords === undefined ||
    maxRecordWords < 3 ||
    maxRecordWords * 2 > bytes.byteLength - headerOffset ||
    parameterCount !== 0
  )
    return null;

  const initialExtent = placeableExtent;
  if (
    initialExtent &&
    (initialExtent.x === 0 ||
      initialExtent.y === 0 ||
      Math.abs(initialExtent.x) > MAX_WMF_DIMENSION ||
      Math.abs(initialExtent.y) > MAX_WMF_DIMENSION)
  )
    return null;
  let state: State = {
    mapMode: MAP_MODE_TEXT,
    origin: placeableOrigin ?? { x: 0, y: 0 },
    extent: initialExtent,
    viewportOrigin: { x: 0, y: 0 },
    viewportExtent: initialExtent
      ? { x: Math.abs(initialExtent.x), y: Math.abs(initialExtent.y) }
      : null,
    current: { x: 0, y: 0 },
    fillRule: "evenodd",
    pen: { type: "pen", stroke: "#000000", width: 1 },
    brush: { type: "brush", fill: "#ffffff" },
  };
  let outputExtent = state.viewportExtent;
  const stack: State[] = [];
  const objects: Array<GraphicsObject | null> = [];
  const output: string[] = [];
  let outputCharacters = 0;
  let points = 0;
  let records = 0;
  let offset = headerOffset + HEADER_WORDS * 2;
  let sawEof = false;

  const append = (element: string): boolean => {
    outputCharacters += element.length;
    if (outputCharacters > MAX_WMF_OUTPUT_CHARACTERS) return false;
    output.push(element);
    return true;
  };
  const addPoints = (count: number): boolean => {
    if (count <= 0 || count > MAX_WMF_POINTS - points) return false;
    points += count;
    return true;
  };
  const addObject = (object: GraphicsObject): boolean => {
    let handle = objects.indexOf(null);
    if (handle < 0) handle = objects.length;
    if (handle >= objectCount || handle >= MAX_WMF_OBJECTS) return false;
    objects[handle] = object;
    return true;
  };

  while (offset + 6 <= bytes.byteLength) {
    records += 1;
    if (records > MAX_WMF_RECORDS) return null;
    const sizeWords = dword(view, offset, bytes.byteLength);
    const fn = word(view, offset + 4, bytes.byteLength);
    if (sizeWords === undefined || fn === undefined || sizeWords < 3) return null;
    const size = sizeWords * 2;
    if (size > bytes.byteLength - offset || sizeWords > maxRecordWords) return null;
    const end = offset + size;
    const params = offset + 6;

    switch (fn) {
      case META.eof:
        if (size !== 6 || end !== bytes.byteLength) return null;
        sawEof = true;
        break;
      case META.setMapMode: {
        const mode = short(view, params, end);
        if (
          size !== 8 ||
          (mode !== MAP_MODE_TEXT && mode !== MAP_MODE_ISOTROPIC && mode !== MAP_MODE_ANISOTROPIC)
        )
          return null;
        state.mapMode = mode;
        break;
      }
      case META.setBkMode: {
        const mode = word(view, params, end);
        if (size !== 8 || (mode !== 1 && mode !== 2)) return null;
        break;
      }
      case META.setPolyFillMode: {
        const mode = word(view, params, end);
        if (size !== 8 || (mode !== POLY_FILL_ALTERNATE && mode !== POLY_FILL_WINDING)) return null;
        state.fillRule = mode === POLY_FILL_ALTERNATE ? "evenodd" : "nonzero";
        break;
      }
      case META.setWindowOrg:
      case META.setWindowExt:
      case META.setViewportOrg:
      case META.setViewportExt: {
        const y = short(view, params, end);
        const x = short(view, params + 2, end);
        if (size !== 10 || x === undefined || y === undefined) return null;
        if (fn === META.setWindowOrg) state.origin = { x, y };
        else if (fn === META.setViewportOrg) state.viewportOrigin = { x, y };
        else {
          if (
            x === 0 ||
            y === 0 ||
            Math.abs(x) > MAX_WMF_DIMENSION ||
            Math.abs(y) > MAX_WMF_DIMENSION
          )
            return null;
          if (fn === META.setWindowExt) {
            state.extent = { x, y };
            state.viewportExtent ??= { x: Math.abs(x), y: Math.abs(y) };
          } else {
            state.viewportExtent = { x, y };
          }
          outputExtent ??= { x: Math.abs(x), y: Math.abs(y) };
        }
        break;
      }
      case META.saveDc:
        if (size !== 6 || stack.length >= MAX_WMF_STATE_DEPTH) return null;
        stack.push(cloneState(state));
        break;
      case META.restoreDc: {
        const relative = short(view, params, end);
        if (size !== 8 || relative === undefined || relative >= 0 || -relative > stack.length)
          return null;
        let restored: State | undefined;
        for (let index = 0; index < -relative; index += 1) restored = stack.pop();
        if (!restored) return null;
        state = restored;
        break;
      }
      case META.createPenIndirect: {
        const style = word(view, params, end);
        const width = short(view, params + 2, end);
        const widthY = short(view, params + 4, end);
        const color = dword(view, params + 6, end);
        if (
          size !== 16 ||
          style === undefined ||
          width === undefined ||
          widthY === undefined ||
          color === undefined ||
          widthY !== 0 ||
          (style & 0x0f) > 5
        )
          return null;
        if (
          !addObject({
            type: "pen",
            stroke: (style & 0x0f) === 5 ? null : metafileColorRefToCss(color),
            width: Math.abs(width),
          })
        )
          return null;
        break;
      }
      case META.createBrushIndirect: {
        const style = word(view, params, end);
        const color = dword(view, params + 2, end);
        const hatch = word(view, params + 6, end);
        if (
          size !== 14 ||
          style === undefined ||
          color === undefined ||
          hatch === undefined ||
          (style !== 0 && style !== 1)
        )
          return null;
        if (!addObject({ type: "brush", fill: style === 1 ? null : metafileColorRefToCss(color) }))
          return null;
        break;
      }
      case META.selectObject: {
        const handle = word(view, params, end);
        const object = handle === undefined ? undefined : objects[handle];
        if (size !== 8 || handle === undefined || !object) return null;
        if (object.type === "pen") state.pen = object;
        else state.brush = object;
        break;
      }
      case META.deleteObject: {
        const handle = word(view, params, end);
        const object = handle === undefined ? undefined : objects[handle];
        if (size !== 8 || handle === undefined || !object) return null;
        if (state.pen === object) state.pen = null;
        if (state.brush === object) state.brush = null;
        for (const saved of stack) {
          if (saved.pen === object) saved.pen = null;
          if (saved.brush === object) saved.brush = null;
        }
        objects[handle] = null;
        break;
      }
      case META.moveTo:
      case META.lineTo: {
        const y = short(view, params, end);
        const x = short(view, params + 2, end);
        if (size !== 10 || x === undefined || y === undefined) return null;
        if (fn === META.lineTo) {
          const start = mappedPoint(state, state.current);
          const finish = mappedPoint(state, { x, y });
          const stroke = state.pen?.stroke;
          const strokeWidth = mappedPenWidth(state, state.pen);
          if (
            !start ||
            !finish ||
            !stroke ||
            strokeWidth === undefined ||
            !append(
              `<path d="${pathPoint("M", start)} ${pathPoint("L", finish)}" fill="none" stroke="${stroke}" stroke-width="${metafileNumberText(strokeWidth)}"/>`,
            )
          )
            return null;
        }
        state.current = { x, y };
        break;
      }
      case META.polygon:
      case META.polyline: {
        const count = word(view, params, end);
        if (count === undefined || size !== 8 + count * 4 || !addPoints(count)) return null;
        const commands: string[] = [];
        for (let index = 0; index < count; index += 1) {
          const point = readPoint(view, params + 2 + index * 4, end);
          const mapped = point && mappedPoint(state, point);
          if (!mapped) return null;
          commands.push(pathPoint(index === 0 ? "M" : "L", mapped));
        }
        if (fn === META.polygon) commands.push("Z");
        const attributes = paint(state);
        if (!attributes) return null;
        let shapeAttributes = attributes;
        if (fn === META.polyline) {
          const pen = state.pen;
          if (!pen?.stroke) return null;
          const strokeWidth = mappedPenWidth(state, pen);
          if (strokeWidth === undefined) return null;
          shapeAttributes = `fill="none" stroke="${pen.stroke}" stroke-width="${metafileNumberText(strokeWidth)}"`;
        } else {
          shapeAttributes += ` fill-rule="${state.fillRule}"`;
        }
        if (!append(`<path d="${commands.join(" ")}" ${shapeAttributes}/>`)) return null;
        break;
      }
      case META.polyPolygon: {
        const polygonCount = word(view, params, end);
        if (polygonCount === undefined || polygonCount === 0) return null;
        const pointsStart = params + 2 + polygonCount * 2;
        if (pointsStart > end) return null;
        let total = 0;
        const counts: number[] = [];
        for (let index = 0; index < polygonCount; index += 1) {
          const count = word(view, params + 2 + index * 2, end);
          if (count === undefined || count < 2) return null;
          counts.push(count);
          total += count;
        }
        if (size !== 8 + polygonCount * 2 + total * 4 || !addPoints(total)) return null;
        const commands: string[] = [];
        let consumed = 0;
        for (const count of counts) {
          for (let index = 0; index < count; index += 1) {
            const point = readPoint(view, pointsStart + (consumed + index) * 4, end);
            const mapped = point && mappedPoint(state, point);
            if (!mapped) return null;
            commands.push(pathPoint(index === 0 ? "M" : "L", mapped));
          }
          commands.push("Z");
          consumed += count;
        }
        const attributes = paint(state);
        if (
          !attributes ||
          !append(`<path d="${commands.join(" ")}" ${attributes} fill-rule="${state.fillRule}"/>`)
        )
          return null;
        break;
      }
      case META.rectangle:
      case META.ellipse: {
        const bottom = short(view, params, end);
        const right = short(view, params + 2, end);
        const top = short(view, params + 4, end);
        const left = short(view, params + 6, end);
        const first =
          left === undefined || top === undefined
            ? undefined
            : mappedPoint(state, { x: left, y: top });
        const second =
          right === undefined || bottom === undefined
            ? undefined
            : mappedPoint(state, { x: right, y: bottom });
        const attributes = paint(state);
        if (size !== 14 || !first || !second || !attributes) return null;
        const x = Math.min(first.x, second.x);
        const y = Math.min(first.y, second.y);
        const width = Math.abs(second.x - first.x);
        const height = Math.abs(second.y - first.y);
        if (width === 0 || height === 0) return null;
        const element =
          fn === META.rectangle
            ? `<rect x="${metafileNumberText(x)}" y="${metafileNumberText(y)}" width="${metafileNumberText(width)}" height="${metafileNumberText(height)}" ${attributes}/>`
            : `<ellipse cx="${metafileNumberText(x + width / 2)}" cy="${metafileNumberText(y + height / 2)}" rx="${metafileNumberText(width / 2)}" ry="${metafileNumberText(height / 2)}" ${attributes}/>`;
        if (!append(element)) return null;
        break;
      }
      default:
        return null;
    }
    offset = end;
    if (sawEof) break;
  }

  if (!sawEof || !outputExtent || output.length === 0) return null;
  const width = Math.abs(outputExtent.x);
  const height = Math.abs(outputExtent.y);
  if (width <= 0 || height <= 0 || width > MAX_WMF_DIMENSION || height > MAX_WMF_DIMENSION)
    return null;
  return assembleMetafileSvg({
    width,
    height,
    elements: output,
    maxCharacters: MAX_WMF_OUTPUT_CHARACTERS,
  });
}
