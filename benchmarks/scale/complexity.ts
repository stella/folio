/** JSON-safe result from comparing CPU time across three input sizes. */
export type ComplexityAssessment = {
  readonly status: "pass" | "fail";
  readonly sizes: readonly [number, number, number];
  readonly cpuMs: readonly [number, number, number];
  /** CPU-time growth from each size to the next. */
  readonly timeRatios: readonly [number, number];
  /** CPU-time-per-input growth from each size to the next. */
  readonly normalizedRatios: readonly [number, number];
  readonly threshold: number;
  readonly explanation: string;
};

export type ComplexityOptions = {
  /** Allowed growth in CPU time per input step, accommodating measurement noise. */
  readonly noiseTolerance?: number;
};

const DEFAULT_NOISE_TOLERANCE = 0.25;

/**
 * Fail only when per-input CPU cost grows beyond the tolerance at both
 * consecutive doublings. Requiring both intervals avoids a single noisy sample
 * creating a superlinear verdict; normalizing by input size also handles fixed
 * setup overhead.
 */
export function assessComplexity(
  sizes: readonly [number, number, number],
  cpuMs: readonly [number, number, number],
  { noiseTolerance = DEFAULT_NOISE_TOLERANCE }: ComplexityOptions = {},
): ComplexityAssessment {
  validateInputs(sizes, cpuMs, noiseTolerance);

  const firstSizeRatio = sizes[1] / sizes[0];
  const secondSizeRatio = sizes[2] / sizes[1];
  const firstTimeRatio = cpuMs[1] / cpuMs[0];
  const secondTimeRatio = cpuMs[2] / cpuMs[1];
  const normalizedRatios = [
    firstTimeRatio / firstSizeRatio,
    secondTimeRatio / secondSizeRatio,
  ] as const;
  const threshold = 1 + noiseTolerance;
  const superlinear = normalizedRatios.every((ratio) => ratio > threshold);

  return {
    status: superlinear ? "fail" : "pass",
    sizes,
    cpuMs,
    timeRatios: [firstTimeRatio, secondTimeRatio],
    normalizedRatios,
    threshold,
    explanation: superlinear
      ? `CPU time per input grew ${format(normalizedRatios[0])}× and then ${format(normalizedRatios[1])}×, exceeding the ${format(threshold)}× limit in both intervals.`
      : `CPU time per input grew ${format(normalizedRatios[0])}× and then ${format(normalizedRatios[1])}×; sustained superlinear growth was not detected (limit ${format(threshold)}× per interval).`,
  };
}

function validateInputs(
  sizes: readonly [number, number, number],
  cpuMs: readonly [number, number, number],
  noiseTolerance: number,
): void {
  if (sizes.some((size) => !Number.isFinite(size) || size <= 0)) {
    throw new RangeError("sizes must be positive finite numbers");
  }
  if (!(sizes[0] < sizes[1] && sizes[1] < sizes[2])) {
    throw new RangeError("sizes must be strictly increasing");
  }
  if (cpuMs.some((time) => !Number.isFinite(time) || time <= 0)) {
    throw new RangeError("CPU times must be positive finite numbers");
  }
  if (!Number.isFinite(noiseTolerance) || noiseTolerance < 0) {
    throw new RangeError("noiseTolerance must be a non-negative finite number");
  }
}

function format(value: number): string {
  return Number(value.toFixed(2)).toString();
}
