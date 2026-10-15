import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { Readable } from "node:stream";
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

test("ZIP input normalization preserves complete text and binary representations", async () => {
  const text = "a".repeat(16 * 1024 - 1) + "👩🏽tail";
  const bytes = encodeUtf8(text);
  const padded = new Uint8Array(bytes.length + 4);
  padded.set(bytes, 2);
  const inputs = {
    text,
    promisedText: Promise.resolve(text),
    array: Array.from(bytes),
    blob: new Blob([bytes]),
    view: new DataView(padded.buffer, 2, bytes.length),
    byteStream: Readable.from([bytes.subarray(0, 16 * 1024), bytes.subarray(16 * 1024)]),
    textStream: Readable.from([text.slice(0, 16 * 1024), text.slice(16 * 1024)]),
    promisedBuffer: Promise.resolve(bytes.buffer),
  };
  for (const input of Object.values(inputs)) {
    const zip = new JSZip();
    writeZipEntry(zip, "entry", input);
    const loaded = await JSZip.loadAsync(await zip.generateAsync({ type: "uint8array" }));
    expect(await loaded.file("entry")?.async("uint8array")).toEqual(bytes);
    expect(await loaded.file("entry")?.async("string")).toBe(text);
  }
});

test("explicit string encodings are normalized before byte flags are forced", async () => {
  const zip = new JSZip();
  writeZipEntry(zip, "binary", "\x00\xff\x80", { binary: true });
  writeZipEntry(zip, "base64", "AP+A", { base64: true });
  const loaded = await JSZip.loadAsync(await zip.generateAsync({ type: "uint8array" }));
  for (const name of ["binary", "base64"])
    expect(await loaded.file(name)?.async("uint8array")).toEqual(Uint8Array.of(0, 255, 128));
});

test("ZIP stream errors propagate to archive generation", async () => {
  const zip = new JSZip();
  const stream = new Readable({
    read() {
      this.destroy(new TypeError("broken stream"));
    },
  });
  writeZipEntry(zip, "entry", stream);
  await expect(zip.generateAsync({ type: "uint8array" })).rejects.toThrow("broken stream");
});
