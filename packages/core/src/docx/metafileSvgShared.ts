/** Helpers shared by the fail-closed EMF and WMF SVG renderers. */

export type MetafilePoint = { x: number; y: number };

type TransformMetafilePointOptions = {
  point: MetafilePoint;
  windowOrigin: MetafilePoint;
  windowExtent: MetafilePoint;
  viewportOrigin: MetafilePoint;
  viewportExtent: MetafilePoint;
  outputOrigin: MetafilePoint;
  maxCoordinate: number;
};

export const transformMetafilePoint = ({
  point,
  windowOrigin,
  windowExtent,
  viewportOrigin,
  viewportExtent,
  outputOrigin,
  maxCoordinate,
}: TransformMetafilePointOptions): MetafilePoint | undefined => {
  if (windowExtent.x === 0 || windowExtent.y === 0) {
    return undefined;
  }
  const x =
    ((point.x - windowOrigin.x) * viewportExtent.x) / windowExtent.x +
    viewportOrigin.x -
    outputOrigin.x;
  const y =
    ((point.y - windowOrigin.y) * viewportExtent.y) / windowExtent.y +
    viewportOrigin.y -
    outputOrigin.y;
  if (!Number.isFinite(x) || !Number.isFinite(y)) {
    return undefined;
  }
  if (Math.abs(x) > maxCoordinate || Math.abs(y) > maxCoordinate) {
    return undefined;
  }
  return { x, y };
};

export const metafileNumberText = (value: number): string => {
  const rounded = Math.round(value * 10_000) / 10_000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
};

export const metafileColorRefToCss = (color: number): string => {
  const red = color & 0xff;
  const green = (color >>> 8) & 0xff;
  const blue = (color >>> 16) & 0xff;
  return `#${red.toString(16).padStart(2, "0")}${green
    .toString(16)
    .padStart(2, "0")}${blue.toString(16).padStart(2, "0")}`;
};

type AssembleMetafileSvgOptions = {
  width: number;
  height: number;
  elements: readonly string[];
  maxCharacters: number;
};

export const assembleMetafileSvg = ({
  width,
  height,
  elements,
  maxCharacters,
}: AssembleMetafileSvgOptions): string | null => {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${metafileNumberText(width)} ${metafileNumberText(height)}">${elements.join("")}</svg>`;
  return svg.length <= maxCharacters ? svg : null;
};
