const MAX_PERCENT = 100;
const HARD_STEP_SLOPE = 1_000;

export type ImageLuminance = {
  brightness?: number | null;
  contrast?: number | null;
};

/** Clamp a finite percentage to the range accepted by DrawingML `a:lum`. */
export const normalizeImageLuminancePercent = (value: number): number =>
  Math.max(-MAX_PERCENT, Math.min(MAX_PERCENT, value));

const finitePercent = (value: number | null | undefined): number =>
  value == null || !Number.isFinite(value) ? 0 : normalizeImageLuminancePercent(value);

const styleNumber = (value: number): string => {
  const rounded = Math.round(value * 1_000_000) / 1_000_000;
  return String(Object.is(rounded, -0) ? 0 : rounded);
};

/**
 * Convert DrawingML brightness and contrast percentages into Word's single
 * linear per-channel transfer, expressed as CSS filters without an
 * intermediate clamp changing the result.
 */
export const imageLuminanceFilter = ({
  brightness,
  contrast,
}: ImageLuminance): string | undefined => {
  const bright = finitePercent(brightness) / MAX_PERCENT;
  const contrastAmount = finitePercent(contrast) / MAX_PERCENT;
  if (bright === 0 && contrastAmount === 0) {
    return undefined;
  }

  const slope =
    contrastAmount >= 0
      ? Math.min(HARD_STEP_SLOPE, 1 / (1 - contrastAmount))
      : Math.max(0, 1 + contrastAmount);
  const pivot = 0.5 - bright / 2;
  const intercept = pivot + bright - slope * pivot;

  if (intercept >= 0) {
    const brightnessAmount = slope + 2 * intercept;
    if (brightnessAmount <= 0) {
      return "brightness(0)";
    }
    return `contrast(${styleNumber(slope / brightnessAmount)}) brightness(${styleNumber(brightnessAmount)})`;
  }

  const contrastFactor = 1 - 2 * intercept;
  return `brightness(${styleNumber(slope / contrastFactor)}) contrast(${styleNumber(contrastFactor)})`;
};
