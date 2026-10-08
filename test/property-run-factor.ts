export const NUM_RUNS_FACTOR_ENV = "PROPERTY_TEST_NUM_RUNS_FACTOR";

const readNumRunsFactor = (raw: string | undefined): number => {
  if (raw === undefined) {
    return 1;
  }
  const parsed = Number(raw);
  // A factor below 1 (or non-numeric) would silently weaken nightly coverage;
  // fall back to the neutral factor instead.
  return Number.isFinite(parsed) && parsed >= 1 ? parsed : 1;
};

export const numRunsFactor = (): number => readNumRunsFactor(process.env[NUM_RUNS_FACTOR_ENV]);
