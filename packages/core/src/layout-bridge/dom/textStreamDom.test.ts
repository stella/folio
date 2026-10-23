import { expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../../test/property-testing";
import { paintedTextOffset } from "./textStreamDom";

test("one-unit model ranges expose only the complete painted text's endpoints", () => {
  fc.assert(
    fc.property(
      fc.nat({ max: 1000 }),
      fc.string({ maxLength: 200 }),
      (pmStart, text) => {
        const projection = {
          pmStart,
          pmEnd: pmStart + 1,
          textLength: text.length,
        };
        expect(paintedTextOffset({ ...projection, pmPosition: pmStart })).toBe(
          0,
        );
        expect(
          paintedTextOffset({ ...projection, pmPosition: pmStart + 1 }),
        ).toBe(text.length);
        expect(
          paintedTextOffset({ ...projection, pmPosition: pmStart - 1 }),
        ).toBe(0);
        expect(
          paintedTextOffset({ ...projection, pmPosition: pmStart + 2 }),
        ).toBe(text.length);
      },
    ),
    propertyConfig({ numRuns: 100 }),
  );
});

test("literal model ranges retain their exact UTF-16 offsets", () => {
  fc.assert(
    fc.property(
      fc.record({
        pmStart: fc.nat({ max: 1000 }),
        textLength: fc.integer({ min: 2, max: 200 }),
        offset: fc.integer({ min: -20, max: 220 }),
      }),
      ({ pmStart, textLength, offset }) => {
        expect(
          paintedTextOffset({
            pmStart,
            pmEnd: pmStart + textLength,
            pmPosition: pmStart + offset,
            textLength,
          }),
        ).toBe(Math.max(0, Math.min(textLength, offset)));
      },
    ),
    propertyConfig({ numRuns: 100 }),
  );
});
