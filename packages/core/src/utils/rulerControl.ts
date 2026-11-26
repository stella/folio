import { TWIPS_PER_INCH } from "./units";

// Match the centimetre ruler ticks (rounded OOXML twips).
const TWIPS_PER_CM = 567;

export const HORIZONTAL_RULER_HEIGHT_PX = 22;

const MIN_CONTENT_SIZE_TWIPS = 720;

type RulerBounds = { min: number; max: number };

type HorizontalRulerBoundsOptions = {
  pageWidth: number;
  leftMargin: number;
  rightMargin: number;
  leftIndent: number;
  rightIndent: number;
};

/** Existing pointer bounds, shared with the keyboard controls. */
export function getHorizontalRulerBounds({
  pageWidth,
  leftMargin,
  rightMargin,
  leftIndent,
  rightIndent,
}: HorizontalRulerBoundsOptions) {
  const content = pageWidth - leftMargin - rightMargin;
  return {
    leftMargin: {
      min: 0,
      max: pageWidth - rightMargin - MIN_CONTENT_SIZE_TWIPS,
    },
    rightMargin: {
      min: 0,
      max: pageWidth - leftMargin - MIN_CONTENT_SIZE_TWIPS,
    },
    firstLineIndent: {
      min: -leftIndent,
      max: content - leftIndent - rightIndent - MIN_CONTENT_SIZE_TWIPS,
    },
    leftIndent: { min: 0, max: content - rightIndent - MIN_CONTENT_SIZE_TWIPS },
    rightIndent: { min: 0, max: content - leftIndent - MIN_CONTENT_SIZE_TWIPS },
  };
}

type VerticalRulerBoundsOptions = {
  pageHeight: number;
  topMargin: number;
  bottomMargin: number;
};

export function getVerticalRulerBounds({
  pageHeight,
  topMargin,
  bottomMargin,
}: VerticalRulerBoundsOptions) {
  return {
    topMargin: {
      min: 0,
      max: pageHeight - bottomMargin - MIN_CONTENT_SIZE_TWIPS,
    },
    bottomMargin: {
      min: 0,
      max: pageHeight - topMargin - MIN_CONTENT_SIZE_TWIPS,
    },
  };
}

/** Clamp pointer values without expanding the range to authored outliers. */
export function clampRulerValue(
  value: number,
  { min, max }: RulerBounds,
): number {
  return Math.round(Math.max(min, Math.min(value, max)));
}

type RulerControlOptions = {
  value: number;
  min: number;
  max: number;
  unit: "inch" | "cm";
  onChange?: ((value: number) => void) | undefined;
};

/** Shared value and keyboard contract for ruler margin and indent handles. */
export function createRulerControl({
  value,
  min,
  max,
  unit,
  onChange,
}: RulerControlOptions) {
  // Include authored values outside pointer bounds without allowing a keyboard
  // edit farther outside that range. Merely showing the ruler never edits it.
  const minimum = Math.min(min, value);
  const maximum = Math.max(min, max, value);
  const step = unit === "inch" ? TWIPS_PER_INCH / 8 : TWIPS_PER_CM / 10;
  return {
    attributes: {
      role: onChange ? ("slider" as const) : undefined,
      tabIndex: onChange ? 0 : undefined,
      "aria-valuenow": onChange ? value : undefined,
      "aria-valuemin": onChange ? minimum : undefined,
      "aria-valuemax": onChange ? maximum : undefined,
      "aria-valuetext": onChange
        ? `${(value / (unit === "inch" ? TWIPS_PER_INCH : TWIPS_PER_CM)).toFixed(2)} ${unit === "inch" ? "in" : "cm"}`
        : undefined,
    },
    onKeyDown(
      event: Pick<
        KeyboardEvent,
        | "key"
        | "altKey"
        | "ctrlKey"
        | "metaKey"
        | "preventDefault"
        | "stopPropagation"
      >,
    ) {
      if (!onChange || event.altKey || event.ctrlKey || event.metaKey) return;
      let next: number;
      switch (event.key) {
        case "ArrowRight":
        case "ArrowUp":
          next = value + step;
          break;
        case "ArrowLeft":
        case "ArrowDown":
          next = value - step;
          break;
        case "PageUp":
          next = value + step * 10;
          break;
        case "PageDown":
          next = value - step * 10;
          break;
        case "Home":
          next = minimum;
          break;
        case "End":
          next = maximum;
          break;
        default:
          return;
      }
      event.preventDefault();
      event.stopPropagation();
      const clamped = Math.round(Math.max(minimum, Math.min(maximum, next)));
      if (clamped !== value) onChange(clamped);
    },
  };
}
