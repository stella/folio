import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { assertProperty, propertyTestTimeout } from "../../../test/property-testing";
import { createTextReader, encodeUtf8 } from "./encoding";
import { writeZipEntry } from "./zip";
import { truncateCodePoints, truncateGraphemes } from "./sliceText";
import { assertTextRoundTrip, textAtBoundary } from "./testing";

setDefaultTimeout(propertyTestTimeout(30_000));

test(
  "ZIP and streamed decoding preserve text at every boundary",
  async () => {
    await assertProperty(
      fc.asyncProperty(textAtBoundary, async (text) => {
        for (const compression of ["STORE", "DEFLATE"] as const) {
          await assertTextRoundTrip({
            text,
            roundTrip: async (value) => {
              const zip = new JSZip();
              expect(writeZipEntry(zip, "text.xml", value)).toBe(zip);
              const loaded = await JSZip.loadAsync(
                await zip.generateAsync({ type: "uint8array", compression }),
              );
              return (await loaded.file("text.xml")?.async("string")) ?? "";
            },
          });
        }
        const bytes = encodeUtf8(text);
        for (const chunkSize of [7, 4096, 16 * 1024, 64 * 1024]) {
          const reader = createTextReader({ fatal: true, ignoreBOM: true });
          const chunks: string[] = [];
          for (let offset = 0; offset < bytes.length; offset += chunkSize)
            chunks.push(reader.read(bytes.subarray(offset, offset + chunkSize)));
          expect(chunks.join("") + reader.finish()).toBe(text);
        }
      }),
      { numRuns: 20, id: "shared ZIP and stream boundaries retain Unicode" },
    );
  },
  propertyTestTimeout(30_000),
);

test("the round-trip assertion detects raw ZIP string encoding", async () => {
  const text = "a".repeat(16 * 1024 - 1) + "👩🏽tail";
  await expect(
    assertTextRoundTrip({
      text,
      roundTrip: async (value) => {
        const zip = new JSZip();
        zip.file("text.xml", value);
        const loaded = await JSZip.loadAsync(await zip.generateAsync({ type: "uint8array" }));
        return (await loaded.file("text.xml")?.async("string")) ?? "";
      },
    }),
  ).rejects.toThrow("Text round trip changed Unicode content");
});

test("byte arrays and buffers preserve bytes and entry options", async () => {
  const bytes = Uint8Array.of(0, 255, 128, 1);
  const date = new Date(Date.UTC(2020, 0, 1));
  const zip = new JSZip();
  writeZipEntry(zip, "bytes", bytes, { date });
  writeZipEntry(zip, "buffer", bytes.buffer);
  writeZipEntry(zip, "promised", Promise.resolve(bytes));
  expect(zip.file("bytes")?.date).toEqual(date);
  const loaded = await JSZip.loadAsync(await zip.generateAsync({ type: "uint8array" }));
  for (const name of ["bytes", "buffer", "promised"])
    expect(await loaded.file(name)?.async("uint8array")).toEqual(bytes);
});

test("truncation retains complete scalars and graphemes", async () => {
  await assertProperty(
    fc.property(textAtBoundary, fc.nat({ max: 30 }), (text, limit) => {
      expect(truncateCodePoints(text, limit)).toBe(Array.from(text).slice(0, limit).join(""));
      const segmenter = new Intl.Segmenter(undefined, { granularity: "grapheme" });
      const graphemes = Array.from(segmenter.segment(text), ({ segment }) => segment);
      expect(truncateGraphemes(text, limit)).toBe(graphemes.slice(0, limit).join(""));
    }),
    { numRuns: 20, id: "shared truncation retains complete text units" },
  );
  expect(truncateGraphemes("👩🏽‍⚖️e\u0301שלום", 1)).toBe("👩🏽‍⚖️");
  expect(truncateCodePoints("😀e\u0301", 1)).toBe("😀");
  expect(truncateCodePoints("abc", -1)).toBe("");
});
