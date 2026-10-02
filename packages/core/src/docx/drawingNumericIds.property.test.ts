import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { relationshipIdOf } from "@stll/docx-core/model";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Image, Shape } from "../types/document";
import { resetAutoIdCounter, serializeRun } from "./serializer/runSerializer";

setDefaultTimeout(propertyTestTimeout(10_000));

const drawingModelId = fc.oneof(
  fc.string(),
  fc.integer({ min: 0, max: 0xffff_ffff }).map(String),
  fc.constantFrom("bare", "_x0000_s1025", "4294967296", "-1", "1.5", ""),
  fc.constant(undefined),
);

// The range oracle must include zero: unsigned DrawingML IDs need no replacement there.
test("every rebuilt drawing uses bounded numeric ids without changing its lexical model id", () => {
  assertProperty(
    fc.property(drawingModelId, (generatedId) => {
      for (const id of [generatedId, "0", "+000", "-0", "4294967295"]) {
        resetAutoIdCounter();
        const image = {
          type: "image",
          id,
          rId: relationshipIdOf("rId1"),
          size: { width: 914_400, height: 914_400 },
          wrap: { type: "inline" },
        } as const satisfies Image;
        const shape = {
          type: "shape",
          id,
          shapeType: "rect",
          shapeNames: { name: "Shape" },
          size: { width: 914_400, height: 914_400 },
        } as const satisfies Shape;
        const xml = serializeRun({
          type: "run",
          content: [
            { type: "drawing", image },
            { type: "shape", shape },
          ],
        });
        const identifiers = [...xml.matchAll(/<(?:wp:docPr|pic:cNvPr|wps:cNvPr) id="([^"]*)"/gu)];
        expect(identifiers).toHaveLength(4);
        const numericSource = id === undefined ? NaN : Number(id);
        const sourceIsNumeric =
          id !== undefined &&
          /^[+-]?\d+$/u.test(id.trim()) &&
          numericSource >= 0 &&
          numericSource <= 0xffff_ffff;
        for (const [, written] of identifiers) {
          expect(written?.trim()).toMatch(/^[+-]?\d+$/u);
          expect(Number(written)).toBeGreaterThanOrEqual(0);
          expect(Number(written)).toBeLessThanOrEqual(0xffff_ffff);
          if (sourceIsNumeric) expect(written).toBe(id);
        }
        expect(image.id).toBe(id);
        expect(shape.id).toBe(id);
      }
    }),
    { numRuns: 150 },
  );
});
