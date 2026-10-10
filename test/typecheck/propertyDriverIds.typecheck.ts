import fc from "fast-check";
import { assertKnownProperty, assertPinnedProperty, assertProperty } from "../property-testing";

// Compiled by typecheck:tooling, outside production package budgets; never executed.
export const propertyDriverIdsProof = () => {
  const sync = fc.property(fc.integer(), () => true);
  const async = fc.asyncProperty(fc.integer(), () => Promise.resolve(true));
  assertProperty(sync, { id: "sync" });
  assertProperty(async, { id: "async" });
  assertPinnedProperty(sync, { id: "pinned" });
  assertKnownProperty(sync, { id: "known", expectedFailures: [] });
  // @ts-expect-error an options argument with an explicit ID is required
  assertProperty(sync);
  // @ts-expect-error a run budget does not supply a stable ID
  assertProperty(sync, { numRuns: 1 });
  // @ts-expect-error async properties also require an explicit ID
  assertProperty(async, { numRuns: 1 });
  // @ts-expect-error pinned drivers require an explicit ID
  assertPinnedProperty(sync, {});
  // @ts-expect-error known-failure witnesses do not supply a stable ID
  assertKnownProperty(sync, { expectedFailures: [] });
};
