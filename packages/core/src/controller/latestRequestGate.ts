/**
 * Coordinates async work where only the most recently started request may
 * commit a result. Each guard closes over an opaque identity, so callers
 * cannot accidentally make an older request current again.
 */
export type LatestRequestGate = {
  /** Start a request and return a guard that remains true while it is latest. */
  begin: () => () => boolean;
  /** Supersede the current request without starting another one. */
  invalidate: () => void;
};

export const createLatestRequestGate = (): LatestRequestGate => {
  let current = Symbol();

  return {
    begin: () => {
      const request = Symbol();
      current = request;
      return () => current === request;
    },
    invalidate: () => {
      current = Symbol();
    },
  };
};
