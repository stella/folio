/** No generated document, and no operation on one, ends a section in a table cell. */

import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { applyDocumentOp } from "../apply";
import { hasSectionInTableCell } from "../contract";
import {
  documentArbitrary,
  opFor,
  opSeedArbitrary,
  reviewDocumentArbitrary,
  trackedOpFor,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(60_000));

const NUM_RUNS = 1_000;

test("no operation on a generated document ends a section in a table cell", () => {
  assertProperty(
    fc.property(
      fc.oneof(documentArbitrary, reviewDocumentArbitrary),
      opSeedArbitrary,
      (document, seed) => {
        expect(hasSectionInTableCell(document)).toBe(false);
        for (const op of [
          opFor(document, seed),
          trackedOpFor(document, seed),
        ]) {
          const applied = applyDocumentOp(document, op);
          if (applied.isOk())
            expect(hasSectionInTableCell(applied.value.document)).toBe(false);
        }
      },
    ),
    {
      id: "no operation on a generated document ends a section in a table cell",
      numRuns: NUM_RUNS,
    },
  );
});
