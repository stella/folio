/**
 * The replayed markup and the typed projection must read a length the same way.
 *
 * A cell width reaches a saved package twice: as the number the capture writes
 * into the replayed `w:tcPr`, and as `TableMeasurement.value` the layout and
 * the serializer work from. If the two disagreed, a Strict document would lay
 * out at one width and save at another, and no example test would say which
 * unit had drifted.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";

import { captureVerbatimXml } from "./verbatimCapture";
import { parseTableMeasurement } from "./tableParser";
import { parseXml, type XmlElement } from "./xmlParser";

const STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const UNITS = ["mm", "cm", "in", "pt", "pc", "pi"] as const;

const cellWidth = (value: string): XmlElement => {
  const root = parseXml(`<w:tcW xmlns:w="${STRICT_W}" w:w="${value}" w:type="dxa"/>`).elements?.[0];
  if (root === undefined) {
    throw new Error("fixture did not parse");
  }
  return root;
};

describe("captured and projected lengths agree", () => {
  test("a universal measure reaches the same twips through both paths", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: 0, max: 100_000 }),
        fc.integer({ min: 0, max: 99 }),
        fc.constantFrom(...UNITS),
        (whole, hundredths, unit) => {
          const element = cellWidth(`${whole}.${String(hundredths).padStart(2, "0")}${unit}`);

          const captured = /w:w="(-?\d+)"/u.exec(captureVerbatimXml(element))?.[1];
          const projected = parseTableMeasurement(element)?.value;

          expect(captured).toBeDefined();
          expect(Number(captured)).toBe(projected ?? Number.NaN);
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });
});
