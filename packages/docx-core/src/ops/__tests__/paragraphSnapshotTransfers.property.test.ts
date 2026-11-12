/**
 * The paragraph property snapshots an edit carries over, read from the
 * top-level blocks it changed, are the ones read from every paragraph of the
 * stories it changed: the same originals for the same replacements, in the
 * same order. Runs of generated operations over generated documents compare
 * them after every edit that applies.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { applyDocumentOp } from "../apply";
import {
  paragraphSnapshotTransfers,
  paragraphSnapshotTransfersReadingEveryParagraph,
} from "../paragraphSnapshotTransfers";
import {
  documentArbitrary,
  opFor,
  opSeedArbitrary,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

/** The property's run count, which the thresholds were set for. */
const COVERAGE_RUNS = 300;

describe("an edit's paragraph property snapshot transfers", () => {
  test("read from the changed blocks equal those read from every paragraph", () => {
    const law = (counts: { compared: number; transferred: number }) =>
      fc.property(
        documentArbitrary,
        fc.array(opSeedArbitrary, { minLength: 1, maxLength: 12 }),
        (initial, seeds) => {
          let document = initial;
          for (const seed of seeds) {
            const applied = applyDocumentOp(document, opFor(document, seed));
            if (applied.isErr()) continue;
            const options = { before: document, edit: applied.value };
            const kept = paragraphSnapshotTransfers(options);
            expect(kept).toEqual(
              paragraphSnapshotTransfersReadingEveryParagraph(options),
            );
            // The same records, not only equal ones.
            const every =
              paragraphSnapshotTransfersReadingEveryParagraph(options);
            kept.forEach(({ source, target }, index) => {
              expect(source).toBe(every[index]?.source);
              expect(target).toBe(every[index]?.target);
            });
            counts.compared += 1;
            counts.transferred += kept.length;
            document = applied.value.document;
          }
        },
      );
    assertProperty(law({ compared: 0, transferred: 0 }), {
      id: "paragraph snapshot transfers read from changed blocks equal every paragraph",
      numRuns: 300,
    });
    const counts = { compared: 0, transferred: 0 };
    checkCoverageSample(law(counts), COVERAGE_RUNS);
    expect(counts.compared).toBeGreaterThan(300);
    expect(counts.transferred).toBeGreaterThan(100);
  });
});
