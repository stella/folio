import { test } from "bun:test";
import fc from "fast-check";

import { assertProperty } from "../../test/property-testing";
import { limit } from "./helper";

// Scratch proof only: a generated failure in a file the change adds.
test("scratch added property", () => {
  assertProperty(
    fc.property(fc.nat(), (value) => value < limit),
    { id: "scratch added property", numRuns: 200 },
  );
});
