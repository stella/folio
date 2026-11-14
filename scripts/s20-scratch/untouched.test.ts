import { test } from "bun:test";
import fc from "fast-check";

import { assertProperty } from "../../test/property-testing";
import { limit } from "./helper";

// Scratch proof only: a generated failure in a file the change leaves alone.
test("scratch untouched property", () => {
  assertProperty(
    fc.property(fc.nat(), (value) => value < limit),
    { id: "scratch untouched property", numRuns: 200 },
  );
});
