/** Shared by timer producers and browser diagnostics, regardless of bundling. */
export const CANONICAL_INPUT_TIMER_OWNER = "folio.canonical-input-timer";

/** Mark the scheduled wrapper, never the caller's potentially shared callback. */
export const setCanonicalInputTimer = (callback: () => void, delay: number) => {
  const ownedCallback = () => callback();
  Object.defineProperty(ownedCallback, Symbol.for(CANONICAL_INPUT_TIMER_OWNER), { value: true });
  return setTimeout(ownedCallback, delay);
};
