import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { encodeUtf8 } from "@stll/utf8-boundaries";
import { writeZipPart } from "./writeZipPart";

setDefaultTimeout(propertyTestTimeout(30_000));

test("nested part writes create exactly their named entries with reproducible dates", async () => {
  const names = ["word/document.xml", "docProps/core.xml", "word/media/nested/image.png"];
  const first = new JSZip();
  const second = new JSZip();
  for (const zip of [first, second]) {
    for (const path of names) writeZipPart({ zip, path, data: encodeUtf8(path) });
    expect(Object.keys(zip.files)).toEqual(names);
    for (const file of Object.values(zip.files)) {
      expect(file.date).toEqual(new Date(Date.UTC(2000, 0, 1)));
      expect(file.dir).toBe(false);
    }
  }
  expect(await first.generateAsync({ type: "uint8array" })).toEqual(
    await second.generateAsync({ type: "uint8array" }),
  );
});

test("explicit entry metadata is preserved without creating parents", () => {
  const zip = new JSZip();
  const date = new Date(Date.UTC(2021, 3, 5));
  writeZipPart({
    zip,
    path: "word/document.xml",
    data: encodeUtf8("<document/>"),
    options: { date, compression: "DEFLATE", compressionOptions: { level: 6 } },
  });
  expect(Object.keys(zip.files)).toEqual(["word/document.xml"]);
  expect(zip.file("word/document.xml")?.date).toEqual(date);
});

const ZIP_STRING_CHUNK_SIZE = 16 * 1024;

test("text parts preserve supplementary characters across archive chunk boundaries", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.record({
        chunk: fc.integer({ min: 1, max: 4 }),
        codePoint: fc.integer({ min: 0x10000, max: 0x10ffff }),
      }),
      async ({ chunk, codePoint }) => {
        for (const compression of ["STORE", "DEFLATE"] as const) {
          for (const shift of [-1, 0, 1]) {
            const zip = new JSZip();
            const expected = new Map<string, string>();
            for (let index = 1; index <= chunk; index++) {
              const boundary = index * ZIP_STRING_CHUNK_SIZE + shift;
              const text = `${"a".repeat(boundary)}${String.fromCodePoint(codePoint)}end`;
              const parts = [
                { path: `word/text-${index}.xml`, data: encodeUtf8(text) },
                { path: `word/promised-${index}.xml`, data: Promise.resolve(encodeUtf8(text)) },
                { path: `word/duplicate-${index}.xml`, data: encodeUtf8(text) },
              ];
              for (const part of parts) {
                writeZipPart({ zip, ...part });
                expected.set(part.path, text);
              }
            }
            const archive = await zip.generateAsync({ type: "uint8array", compression });
            const reopened = await JSZip.loadAsync(archive);
            expect(Object.keys(reopened.files).sort()).toEqual([...expected.keys()].sort());
            for (const [path, text] of expected) {
              const bytes = await reopened.file(path)?.async("uint8array");
              const restored = await reopened.file(path)?.async("string");
              expect(bytes).toEqual(new TextEncoder().encode(text));
              expect(restored?.slice(-5)).toBe(text.slice(-5));
              expect(restored === text).toBe(true);
            }
          }
        }
      },
    ),
    { numRuns: 10, id: "ZIP text preserves supplementary characters across chunk boundaries" },
  );
});

test("byte inputs retain their values", async () => {
  const bytes = Uint8Array.of(0xff, 0, 0x80);
  const zip = new JSZip();
  writeZipPart({ zip, path: "bytes", data: bytes });
  writeZipPart({ zip, path: "promised-bytes", data: Promise.resolve(bytes) });
  const reopened = await JSZip.loadAsync(await zip.generateAsync({ type: "uint8array" }));
  for (const name of ["bytes", "promised-bytes"]) {
    expect(await reopened.file(name)?.async("uint8array")).toEqual(bytes);
  }
});

test("directory entries ignore payload encoding", async () => {
  for (const options of [
    { dir: true },
    { unixPermissions: 0o40755 },
    { unixPermissions: "40755" },
    { dosPermissions: 0x10 },
  ]) {
    const zip = new JSZip();
    writeZipPart({
      zip,
      path: "folder",
      data: Uint8Array.of(1, 2, 3),
      options,
    });
    expect(Object.keys(zip.files)).toEqual(["folder/"]);
    expect(await zip.files["folder/"]?.async("uint8array")).toEqual(new Uint8Array());
  }
});
