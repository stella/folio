/**
 * Render a deliberately small, browser-safe EMF subset to generated SVG.
 *
 * This covers vector logos built from filled paths. Unsupported records fail
 * closed: a partial metafile is more misleading than the existing placeholder.
 * Original EMF bytes remain the source of truth for DOCX round-trip.
 */

const EMF_SIGNATURE = 0x464d_4520;
const MAX_EMF_RECORDS = 10_000;
const MAX_EMF_POINTS = 20_000;
const MAX_EMF_DIMENSION = 100_000;
const MAX_EMF_SVG_CHARACTERS = 1_000_000;
const MAX_EMF_STATE_DEPTH = 256;

const EMR = {
  header: 1,
  setWindowExtEx: 9,
  setWindowOrgEx: 10,
  setViewportExtEx: 11,
  setViewportOrgEx: 12,
  eof: 14,
  setMapMode: 17,
  setBackgroundMode: 18,
  setPolyFillMode: 19,
  moveToEx: 27,
  saveDc: 33,
  restoreDc: 34,
  selectObject: 37,
  createBrushIndirect: 39,
  deleteObject: 40,
  beginPath: 59,
  endPath: 60,
  closeFigure: 61,
  fillPath: 62,
  polyBezierTo16: 88,
  polyPolygon16: 91,
  extCreatePen: 95,
} as const;

const MAP_MODE_TEXT = 1;
const MAP_MODE_ANISOTROPIC = 8;
const POLY_FILL_ALTERNATE = 1;
const POLY_FILL_WINDING = 2;
const STOCK_OBJECT_FLAG = 0x8000_0000;

type Brush = { type: "brush"; fill: string | null };
type Pen = { type: "pen"; stroke: null };
type GraphicsObject = Brush | Pen;

type PathState =
  | { type: "idle" }
  | { type: "building"; commands: string[] }
  | { type: "completed"; commands: string[] };

type DeviceState = {
  mapMode: typeof MAP_MODE_TEXT | typeof MAP_MODE_ANISOTROPIC;
  windowOrigin: { x: number; y: number };
  windowExtent: { x: number; y: number };
  viewportOrigin: { x: number; y: number };
  viewportExtent: { x: number; y: number };
  fillRule: "evenodd" | "nonzero";
  brush: Brush | null;
  pen: Pen | null;
  currentPoint: { x: number; y: number };
};

const initialState = (): DeviceState => ({
  mapMode: MAP_MODE_TEXT,
  windowOrigin: { x: 0, y: 0 },
  windowExtent: { x: 1, y: 1 },
  viewportOrigin: { x: 0, y: 0 },
  viewportExtent: { x: 1, y: 1 },
  fillRule: "nonzero",
  brush: null,
  pen: null,
  currentPoint: { x: 0, y: 0 },
});

const cloneState = (state: DeviceState): DeviceState => ({
  ...state,
  windowOrigin: { ...state.windowOrigin },
  windowExtent: { ...state.windowExtent },
  viewportOrigin: { ...state.viewportOrigin },
  viewportExtent: { ...state.viewportExtent },
  currentPoint: { ...state.currentPoint },
});

const int16 = (view: DataView, offset: number, end: number): number | undefined =>
  offset >= 0 && offset + 2 <= end ? view.getInt16(offset, true) : undefined;

const uint32 = (view: DataView, offset: number, end: number): number | undefined =>
  offset >= 0 && offset + 4 <= end ? view.getUint32(offset, true) : undefined;

const int32 = (view: DataView, offset: number, end: number): number | undefined =>
  offset >= 0 && offset + 4 <= end ? view.getInt32(offset, true) : undefined;

const numberText = (value: number): string => {
  const rounded = Math.round(value * 10_000) / 10_000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
};

const colorRefToCss = (color: number): string => {
  const red = color & 0xff;
  const green = (color >>> 8) & 0xff;
  const blue = (color >>> 16) & 0xff;
  return `#${red.toString(16).padStart(2, "0")}${green
    .toString(16)
    .padStart(2, "0")}${blue.toString(16).padStart(2, "0")}`;
};

const stockObject = (handle: number): GraphicsObject | undefined => {
  switch (handle) {
    case 0x8000_0000:
      return { type: "brush", fill: "#ffffff" };
    case 0x8000_0001:
      return { type: "brush", fill: "#c0c0c0" };
    case 0x8000_0002:
      return { type: "brush", fill: "#808080" };
    case 0x8000_0003:
      return { type: "brush", fill: "#404040" };
    case 0x8000_0004:
      return { type: "brush", fill: "#000000" };
    case 0x8000_0005:
      return { type: "brush", fill: null };
    case 0x8000_0008:
      return { type: "pen", stroke: null };
    default:
      return undefined;
  }
};

type EmfBounds = { left: number; top: number; width: number; height: number };

const mappedPoint = (
  state: DeviceState,
  bounds: EmfBounds,
  x: number,
  y: number,
): { x: number; y: number } | undefined => {
  let mappedX: number;
  let mappedY: number;
  if (state.mapMode === MAP_MODE_TEXT) {
    mappedX = x - state.windowOrigin.x + state.viewportOrigin.x - bounds.left;
    mappedY = y - state.windowOrigin.y + state.viewportOrigin.y - bounds.top;
  } else {
    const { windowExtent, windowOrigin, viewportExtent, viewportOrigin } = state;
    if (windowExtent.x === 0 || windowExtent.y === 0) {
      return undefined;
    }
    mappedX =
      ((x - windowOrigin.x) * viewportExtent.x) / windowExtent.x + viewportOrigin.x - bounds.left;
    mappedY =
      ((y - windowOrigin.y) * viewportExtent.y) / windowExtent.y + viewportOrigin.y - bounds.top;
  }
  if (!Number.isFinite(mappedX) || !Number.isFinite(mappedY)) {
    return undefined;
  }
  if (Math.abs(mappedX) > MAX_EMF_DIMENSION * 100 || Math.abs(mappedY) > MAX_EMF_DIMENSION * 100) {
    return undefined;
  }
  return { x: mappedX, y: mappedY };
};

const pathPoint = (command: string, point: { x: number; y: number }): string =>
  `${command}${numberText(point.x)} ${numberText(point.y)}`;

const selectGraphicsObject = (
  state: DeviceState,
  objects: Map<number, GraphicsObject>,
  handle: number,
): boolean => {
  const object = handle >= STOCK_OBJECT_FLAG ? stockObject(handle) : objects.get(handle);
  if (!object) {
    return false;
  }
  if (object.type === "brush") {
    state.brush = object;
  } else {
    state.pen = object;
  }
  return true;
};

/** Convert the supported filled-path EMF subset to inert generated SVG. */
export function renderEmfSvg(data: ArrayBuffer | Uint8Array): string | null {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (bytes.byteLength < 88) {
    return null;
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const headerType = uint32(view, 0, bytes.byteLength);
  const headerSize = uint32(view, 4, bytes.byteLength);
  const signature = uint32(view, 40, bytes.byteLength);
  const declaredBytes = uint32(view, 48, bytes.byteLength);
  const declaredRecords = uint32(view, 52, bytes.byteLength);
  const left = int32(view, 8, bytes.byteLength);
  const top = int32(view, 12, bytes.byteLength);
  const right = int32(view, 16, bytes.byteLength);
  const bottom = int32(view, 20, bytes.byteLength);
  if (
    headerType !== EMR.header ||
    headerSize === undefined ||
    headerSize < 88 ||
    headerSize % 4 !== 0 ||
    headerSize > bytes.byteLength ||
    signature !== EMF_SIGNATURE ||
    declaredBytes !== bytes.byteLength ||
    declaredRecords === undefined ||
    declaredRecords < 2 ||
    declaredRecords > MAX_EMF_RECORDS ||
    left === undefined ||
    top === undefined ||
    right === undefined ||
    bottom === undefined
  ) {
    return null;
  }
  const width = right - left;
  const height = bottom - top;
  if (width <= 0 || height <= 0 || width > MAX_EMF_DIMENSION || height > MAX_EMF_DIMENSION) {
    return null;
  }

  const bounds = { left, top, width, height };
  const objects = new Map<number, GraphicsObject>();
  const stack: DeviceState[] = [];
  let state = initialState();
  let offset = headerSize;
  let recordCount = 1;
  let pointCount = 0;
  let pathState: PathState = { type: "idle" };
  const renderedPaths: string[] = [];
  let sawEof = false;

  while (offset + 8 <= bytes.byteLength && recordCount < MAX_EMF_RECORDS) {
    const type = uint32(view, offset, bytes.byteLength);
    const size = uint32(view, offset + 4, bytes.byteLength);
    if (
      type === undefined ||
      size === undefined ||
      size < 8 ||
      size % 4 !== 0 ||
      size > bytes.byteLength - offset
    ) {
      return null;
    }
    const end = offset + size;
    recordCount += 1;

    switch (type) {
      case EMR.eof:
        if (size !== 20 || end !== bytes.byteLength || recordCount !== declaredRecords) {
          return null;
        }
        sawEof = true;
        break;
      case EMR.setMapMode: {
        const mode = int32(view, offset + 8, end);
        if (size !== 12 || (mode !== MAP_MODE_TEXT && mode !== MAP_MODE_ANISOTROPIC)) {
          return null;
        }
        state.mapMode = mode;
        break;
      }
      case EMR.setBackgroundMode:
        if (size !== 12 || uint32(view, offset + 8, end) === undefined) {
          return null;
        }
        break;
      case EMR.setPolyFillMode: {
        const mode = uint32(view, offset + 8, end);
        if (size !== 12 || (mode !== POLY_FILL_ALTERNATE && mode !== POLY_FILL_WINDING)) {
          return null;
        }
        state.fillRule = mode === POLY_FILL_ALTERNATE ? "evenodd" : "nonzero";
        break;
      }
      case EMR.setWindowOrgEx:
      case EMR.setWindowExtEx:
      case EMR.setViewportOrgEx:
      case EMR.setViewportExtEx: {
        const x = int32(view, offset + 8, end);
        const y = int32(view, offset + 12, end);
        if (size !== 16 || x === undefined || y === undefined) {
          return null;
        }
        if (type === EMR.setWindowOrgEx) {
          state.windowOrigin = { x, y };
        } else if (type === EMR.setWindowExtEx) {
          if (state.mapMode === MAP_MODE_ANISOTROPIC) {
            if (x === 0 || y === 0) {
              return null;
            }
            state.windowExtent = { x, y };
          }
        } else if (type === EMR.setViewportOrgEx) {
          state.viewportOrigin = { x, y };
        } else if (state.mapMode === MAP_MODE_ANISOTROPIC) {
          if (x === 0 || y === 0) {
            return null;
          }
          state.viewportExtent = { x, y };
        }
        break;
      }
      case EMR.saveDc:
        if (size !== 8 || stack.length >= MAX_EMF_STATE_DEPTH) {
          return null;
        }
        stack.push(cloneState(state));
        break;
      case EMR.restoreDc: {
        const relative = int32(view, offset + 8, end);
        if (size !== 12 || relative === undefined || relative >= 0 || -relative > stack.length) {
          return null;
        }
        let restored: DeviceState | undefined;
        for (let index = 0; index < -relative; index += 1) {
          restored = stack.pop();
        }
        if (!restored) {
          return null;
        }
        state = restored;
        break;
      }
      case EMR.createBrushIndirect: {
        const handle = uint32(view, offset + 8, end);
        const style = uint32(view, offset + 12, end);
        const color = uint32(view, offset + 16, end);
        if (
          size !== 24 ||
          handle === undefined ||
          handle >= STOCK_OBJECT_FLAG ||
          objects.has(handle) ||
          style === undefined ||
          color === undefined
        ) {
          return null;
        }
        if (style !== 0 && style !== 1) {
          return null;
        }
        objects.set(handle, {
          type: "brush",
          fill: style === 1 ? null : colorRefToCss(color),
        });
        break;
      }
      case EMR.extCreatePen: {
        const handle = uint32(view, offset + 8, end);
        const style = uint32(view, offset + 28, end);
        const bmiBytes = uint32(view, offset + 16, end);
        const bitmapBytes = uint32(view, offset + 24, end);
        const styleEntries = uint32(view, offset + 48, end);
        if (
          size !== 52 ||
          handle === undefined ||
          handle >= STOCK_OBJECT_FLAG ||
          objects.has(handle) ||
          style === undefined ||
          bmiBytes !== 0 ||
          bitmapBytes !== 0 ||
          styleEntries !== 0 ||
          (style & 0x0f) !== 5
        ) {
          return null;
        }
        objects.set(handle, { type: "pen", stroke: null });
        break;
      }
      case EMR.selectObject: {
        const handle = uint32(view, offset + 8, end);
        if (size !== 12 || handle === undefined || !selectGraphicsObject(state, objects, handle)) {
          return null;
        }
        break;
      }
      case EMR.deleteObject: {
        const handle = uint32(view, offset + 8, end);
        if (size !== 12 || handle === undefined) {
          return null;
        }
        const object = objects.get(handle);
        if (object === state.brush) {
          state.brush = null;
        } else if (object === state.pen) {
          state.pen = null;
        }
        for (const savedState of stack) {
          if (object === savedState.brush) {
            savedState.brush = null;
          } else if (object === savedState.pen) {
            savedState.pen = null;
          }
        }
        objects.delete(handle);
        break;
      }
      case EMR.beginPath: {
        if (size !== 8) {
          return null;
        }
        switch (pathState.type) {
          case "idle":
            pathState = { type: "building", commands: [] };
            break;
          case "building":
          case "completed":
            return null;
          default:
            pathState satisfies never;
        }
        break;
      }
      case EMR.moveToEx: {
        const x = int32(view, offset + 8, end);
        const y = int32(view, offset + 12, end);
        if (size !== 16 || x === undefined || y === undefined) {
          return null;
        }
        state.currentPoint = { x, y };
        switch (pathState.type) {
          case "idle":
          case "completed":
            break;
          case "building": {
            const mapped = mappedPoint(state, bounds, x, y);
            if (!mapped) {
              return null;
            }
            pathState.commands.push(pathPoint("M", mapped));
            break;
          }
          default:
            pathState satisfies never;
        }
        break;
      }
      case EMR.polyBezierTo16: {
        let commands: string[];
        switch (pathState.type) {
          case "building":
            commands = pathState.commands;
            break;
          case "idle":
          case "completed":
            return null;
          default:
            pathState satisfies never;
            return null;
        }
        const count = uint32(view, offset + 24, end);
        if (
          count === undefined ||
          count === 0 ||
          count % 3 !== 0 ||
          count > MAX_EMF_POINTS - pointCount ||
          size !== 28 + count * 4
        ) {
          return null;
        }
        pointCount += count;
        if (commands.length === 0) {
          const mapped = mappedPoint(state, bounds, state.currentPoint.x, state.currentPoint.y);
          if (!mapped) {
            return null;
          }
          commands.push(pathPoint("M", mapped));
        }
        for (let index = 0; index < count; index += 3) {
          const points: { x: number; y: number }[] = [];
          for (let controlIndex = 0; controlIndex < 3; controlIndex += 1) {
            const pointOffset = offset + 28 + (index + controlIndex) * 4;
            const x = int16(view, pointOffset, end);
            const y = int16(view, pointOffset + 2, end);
            if (x === undefined || y === undefined) {
              return null;
            }
            const mapped = mappedPoint(state, bounds, x, y);
            if (!mapped) {
              return null;
            }
            points.push(mapped);
            state.currentPoint = { x, y };
          }
          const [first, second, third] = points;
          if (!first || !second || !third) {
            return null;
          }
          commands.push(
            `C${numberText(first.x)} ${numberText(first.y)} ${numberText(second.x)} ${numberText(second.y)} ${numberText(third.x)} ${numberText(third.y)}`,
          );
        }
        break;
      }
      case EMR.polyPolygon16: {
        const polygonCount = uint32(view, offset + 24, end);
        const totalPoints = uint32(view, offset + 28, end);
        if (
          polygonCount === undefined ||
          totalPoints === undefined ||
          polygonCount === 0 ||
          totalPoints === 0 ||
          polygonCount > totalPoints ||
          totalPoints > MAX_EMF_POINTS - pointCount ||
          polygonCount > Math.floor((end - (offset + 32)) / 4)
        ) {
          return null;
        }
        const pointsOffset = offset + 32 + polygonCount * 4;
        if (size !== 32 + polygonCount * 4 + totalPoints * 4) {
          return null;
        }
        pointCount += totalPoints;
        let consumedPoints = 0;
        const polygonPath: string[] = [];
        for (let polygonIndex = 0; polygonIndex < polygonCount; polygonIndex += 1) {
          const count = uint32(view, offset + 32 + polygonIndex * 4, end);
          if (count === undefined || count < 2 || count > totalPoints - consumedPoints) {
            return null;
          }
          for (let index = 0; index < count; index += 1) {
            const pointOffset = pointsOffset + (consumedPoints + index) * 4;
            const x = int16(view, pointOffset, end);
            const y = int16(view, pointOffset + 2, end);
            if (x === undefined || y === undefined) {
              return null;
            }
            const mapped = mappedPoint(state, bounds, x, y);
            if (!mapped) {
              return null;
            }
            polygonPath.push(pathPoint(index === 0 ? "M" : "L", mapped));
          }
          polygonPath.push("Z");
          consumedPoints += count;
        }
        if (consumedPoints !== totalPoints) {
          return null;
        }
        switch (pathState.type) {
          case "building":
            pathState.commands.push(...polygonPath);
            break;
          case "idle":
          case "completed":
            if (!state.brush?.fill) {
              return null;
            }
            renderedPaths.push(
              `<path d="${polygonPath.join(" ")}" fill="${state.brush.fill}" fill-rule="${state.fillRule}"/>`,
            );
            break;
          default:
            pathState satisfies never;
        }
        break;
      }
      case EMR.closeFigure: {
        if (size !== 8) {
          return null;
        }
        switch (pathState.type) {
          case "building":
            pathState.commands.push("Z");
            break;
          case "idle":
          case "completed":
            return null;
          default:
            pathState satisfies never;
        }
        break;
      }
      case EMR.endPath: {
        if (size !== 8) {
          return null;
        }
        switch (pathState.type) {
          case "building":
            if (pathState.commands.length === 0) {
              return null;
            }
            pathState = { type: "completed", commands: pathState.commands };
            break;
          case "idle":
          case "completed":
            return null;
          default:
            pathState satisfies never;
        }
        break;
      }
      case EMR.fillPath: {
        if (size !== 24 || !state.brush?.fill) {
          return null;
        }
        switch (pathState.type) {
          case "completed":
            renderedPaths.push(
              `<path d="${pathState.commands.join(" ")}" fill="${state.brush.fill}" fill-rule="${state.fillRule}"/>`,
            );
            pathState = { type: "idle" };
            break;
          case "idle":
          case "building":
            return null;
          default:
            pathState satisfies never;
        }
        break;
      }
      default:
        return null;
    }

    offset = end;
    if (sawEof) {
      break;
    }
  }

  if (
    !sawEof ||
    offset !== bytes.byteLength ||
    recordCount !== declaredRecords ||
    renderedPaths.length === 0
  ) {
    return null;
  }
  switch (pathState.type) {
    case "idle":
      break;
    case "building":
    case "completed":
      return null;
    default:
      pathState satisfies never;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${numberText(width)} ${numberText(height)}">${renderedPaths.join("")}</svg>`;
  return svg.length <= MAX_EMF_SVG_CHARACTERS ? svg : null;
}
