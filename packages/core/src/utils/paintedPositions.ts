/**
 * The one owner of the editor positions a painted element carries.
 *
 * The painter marks the elements it draws with the editor document positions
 * they show: a run, a paragraph line, a table, a table cell, an image, a text
 * box, a content-control boundary. Hit testing, caret and selection drawing,
 * scrolling to a position and the hosts' pointer handling read them back.
 *
 * How a position is stored on an element is this module's business alone:
 * every write and every read goes through it, and the
 * `folio-painted-positions/position-attribute-owner` lint rule forbids
 * touching the attributes anywhere else. That keeps the storage free to
 * change in one place, for instance to positions relative to the painted
 * block, without auditing every reader.
 */

const START_ATTRIBUTE = "data-pm-start";
const END_ATTRIBUTE = "data-pm-end";
const TABLE_START_ATTRIBUTE = "data-table-pm-start";
const CONTENT_CONTROL_ATTRIBUTE = "data-sdt-pm-pos";

/** Elements that show where they start. */
export const PAINTED_START_SELECTOR = `[${START_ATTRIBUTE}]`;

/** Elements that show where they start and where they end. */
export const PAINTED_RANGE_SELECTOR = `[${START_ATTRIBUTE}][${END_ATTRIBUTE}]`;

/** Elements that carry any painted position. */
export const PAINTED_POSITION_SELECTOR = [
  START_ATTRIBUTE,
  END_ATTRIBUTE,
  TABLE_START_ATTRIBUTE,
]
  .map((attribute) => `[${attribute}]`)
  .join(", ");

/** Elements that show a range starting exactly at `position`. */
export const paintedStartSelector = (position: number): string =>
  `[${START_ATTRIBUTE}="${String(position)}"]`;

/**
 * The attributes a node can differ in from a fresh painting of it while still
 * showing the same thing: an edit before it shifts them and nothing else.
 */
export const PAINTED_POSITION_ATTRIBUTES: ReadonlySet<string> = new Set([
  START_ATTRIBUTE,
  END_ATTRIBUTE,
  TABLE_START_ATTRIBUTE,
  CONTENT_CONTROL_ATTRIBUTE,
]);

/**
 * The element side of a painted position: `dataset` reflects the attributes
 * above under these keys.
 */
type PositionKey = "pmStart" | "pmEnd" | "tablePmStart" | "sdtPmPos";

/** An element that can carry painted positions. */
export type PaintedElement = Pick<HTMLElement, "dataset">;

/** A painted value, or `undefined` when the element carries none. */
const readPosition = (
  element: PaintedElement,
  key: PositionKey,
): number | undefined => {
  const value = element.dataset[key];
  return value === undefined || value === "" ? undefined : Number(value);
};

/** Where the content `element` shows starts. */
export const readPaintedStart = (element: PaintedElement): number | undefined =>
  readPosition(element, "pmStart");

/** Where the content `element` shows ends. */
export const readPaintedEnd = (element: PaintedElement): number | undefined =>
  readPosition(element, "pmEnd");

/** Where the table a table handle or overlay element belongs to starts. */
export const readPaintedTableStart = (
  element: PaintedElement,
): number | undefined => readPosition(element, "tablePmStart");

/** The raw content-control boundary key a painted box carries. */
export const readPaintedContentControlKey = (
  element: PaintedElement,
): string | undefined => element.dataset["sdtPmPos"];

/** Whether `element` shows where it starts. */
export const hasPaintedStart = (element: PaintedElement): boolean =>
  readPaintedStart(element) !== undefined;

/** Whether `element` shows where it starts and where it ends. */
export const hasPaintedRange = (element: PaintedElement): boolean =>
  readPaintedStart(element) !== undefined &&
  readPaintedEnd(element) !== undefined;

/** Mark `element` as showing content that starts at `position`. */
export const writePaintedStart = (
  element: PaintedElement,
  position: number,
): void => {
  element.dataset["pmStart"] = String(position);
};

/** Mark `element` as showing content that ends at `position`. */
export const writePaintedEnd = (
  element: PaintedElement,
  position: number,
): void => {
  element.dataset["pmEnd"] = String(position);
};

/**
 * Mark `element` with the range it shows. A side left `undefined` is not
 * written, as for a fragment that knows only one of them.
 */
export const writePaintedRange = (
  element: PaintedElement,
  start: number | undefined,
  end: number | undefined,
): void => {
  if (start !== undefined) writePaintedStart(element, start);
  if (end !== undefined) writePaintedEnd(element, end);
};

/** Mark a table handle or overlay element with its table's start. */
export const writePaintedTableStart = (
  element: PaintedElement,
  position: number,
): void => {
  element.dataset["tablePmStart"] = String(position);
};

/** Mark a content-control box with the boundary it draws. */
export const writePaintedContentControlKey = (
  element: PaintedElement,
  position: number,
): void => {
  element.dataset["sdtPmPos"] = String(position);
};

const shiftPosition = (
  element: PaintedElement,
  key: PositionKey,
  delta: number,
): void => {
  const value = element.dataset[key];
  if (value === undefined) return;
  const position = Number(value);
  if (Number.isFinite(position))
    element.dataset[key] = String(position + delta);
};

/**
 * Move every painted position under `root` (not `root` itself) by `delta`,
 * for content an edit before it moved without changing it.
 */
export const shiftPaintedPositions = (
  root: ParentNode,
  delta: number,
): void => {
  for (const element of root.querySelectorAll<HTMLElement>(
    PAINTED_POSITION_SELECTOR,
  )) {
    shiftPosition(element, "pmStart", delta);
    shiftPosition(element, "pmEnd", delta);
    shiftPosition(element, "tablePmStart", delta);
  }
};
