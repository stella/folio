import { panic } from "better-result";
import { numRunsFactor } from "./property-run-factor";

/** Scale a stated base budget by the same factor as generated case counts. */
export const propertyTestTimeout = (baseMs: number): number => {
  if (!Number.isFinite(baseMs) || baseMs <= 0) {
    return panic("Property test base budget must be finite and positive.", { baseMs });
  }
  const timeout = Math.ceil(baseMs * numRunsFactor());
  if (!Number.isFinite(timeout)) {
    return panic("Scaled property test budget must be finite.", { baseMs });
  }
  return timeout;
};
