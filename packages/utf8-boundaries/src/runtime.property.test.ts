import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../test/property-testing";
import { createTextReader, encodeUtf8 } from "./encoding";
import { truncateCodePoints, truncateGraphemes } from "./sliceText";
import { textAtBoundary } from "./testing";

setDefaultTimeout(propertyTestTimeout(30_000));

test(
  "streamed decoding preserves text at every boundary",
  async () => {
    await assertProperty(
      fc.asyncProperty(textAtBoundary, async (text) => {
        const bytes = encodeUtf8(text);
        for (const chunkSize of [7, 4096, 16 * 1024, 64 * 1024]) {
          const reader = createTextReader({ fatal: true, ignoreBOM: true });
          const chunks: string[] = [];
          for (let offset = 0; offset < bytes.length; offset += chunkSize)
            chunks.push(
              reader.read(bytes.subarray(offset, offset + chunkSize)),
            );
          expect(chunks.join("") + reader.finish()).toBe(text);
        }
      }),
      { numRuns: 20, id: "shared stream boundaries retain Unicode" },
    );
  },
  propertyTestTimeout(30_000),
);

test("truncation retains complete scalars and graphemes", async () => {
  await assertProperty(
    fc.property(textAtBoundary, fc.nat({ max: 30 }), (text, limit) => {
      expect(truncateCodePoints(text, limit)).toBe(
        Array.from(text).slice(0, limit).join(""),
      );
      const segmenter = new Intl.Segmenter(undefined, {
        granularity: "grapheme",
      });
      const graphemes = Array.from(
        segmenter.segment(text),
        ({ segment }) => segment,
      );
      expect(truncateGraphemes(text, limit)).toBe(
        graphemes.slice(0, limit).join(""),
      );
    }),
    { numRuns: 20, id: "shared truncation retains complete text units" },
  );
  expect(truncateGraphemes("👩🏽‍⚖️e\u0301שלום", 1)).toBe("👩🏽‍⚖️");
  expect(truncateCodePoints("😀e\u0301", 1)).toBe("😀");
  expect(truncateCodePoints("abc", -1)).toBe("");
});
