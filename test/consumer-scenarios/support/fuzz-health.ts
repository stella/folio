export type FuzzHealth =
  | { status: "passed"; completed: number }
  | { status: "finding"; completed: number; detail: string }
  | { status: "infrastructure"; completed: number; detail: string };

export const HEALTH_MARKER = "FOLIO_FUZZ_HEALTH ";

/** Enabled only for fuzz jobs; ordinary tests keep their existing output. */
export const reportFuzzHealth = (
  health: FuzzHealth | { status: "started"; completed: 0 },
): void => {
  if (process.env["FOLIO_FUZZ_HEALTH"] !== "1") return;
  console.log(`${HEALTH_MARKER}${JSON.stringify(health)}`);
};
