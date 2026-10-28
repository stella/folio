import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { ZipArchive, ZipError } from "./archive";
import { webZipCodec, type ZipCodec } from "./codec";

const build = async (
  entries: Record<string, string | Uint8Array>,
): Promise<Uint8Array<ArrayBuffer>> => {
  const archive = new ZipArchive();
  for (const [path, data] of Object.entries(entries)) archive.set(path, data);
  return await archive.generate();
};

/** Read every entry; anything thrown that is not a ZipError fails the property. */
const readAllOrZipError = async (
  bytes: Uint8Array,
  codec: ZipCodec | undefined,
): Promise<void> => {
  try {
    const archive = await ZipArchive.read(
      bytes,
      codec === undefined ? {} : { codec },
    );
    for (const entry of archive.entries()) await entry.bytes();
  } catch (error) {
    if (!(error instanceof ZipError)) throw error;
  }
};

describe("arbitrary corruption", () => {
  test(
    "any byte change or truncation either reads or fails with a ZipError",
    async () => {
      const source = await build({
        "[Content_Types].xml": "<Types/>",
        "word/document.xml": "<w:document>".repeat(200),
        "word/media/image.png": Uint8Array.from(
          { length: 300 },
          (_, i) => i % 7,
        ),
      });
      await assertProperty(
        fc.asyncProperty(
          fc.array(
            fc.record({
              offset: fc.nat({ max: source.byteLength - 1 }),
              value: fc.integer({ min: 0, max: 255 }),
            }),
            { minLength: 1, maxLength: 4 },
          ),
          fc.option(fc.nat({ max: source.byteLength }), { nil: undefined }),
          fc.constantFrom<ZipCodec | undefined>(undefined, webZipCodec),
          async (edits, truncate, codec) => {
            const bytes = source.slice(0, truncate ?? source.byteLength);
            for (const { offset, value } of edits) {
              if (offset < bytes.byteLength) bytes[offset] = value;
            }
            await readAllOrZipError(bytes, codec);
          },
        ),
        { numRuns: 200, id: "ZIP corruption is refused with a ZipError" },
      );
    },
    propertyTestTimeout(60_000),
  );

  test(
    "written archives read back entry for entry",
    async () => {
      const segment = fc
        .stringMatching(/^[A-Za-z0-9_.\-\u00e1\u010d]{1,8}$/u)
        .filter((part) => part !== "." && part !== "..");
      await assertProperty(
        fc.asyncProperty(
          fc.uniqueArray(
            fc.record({
              path: fc
                .array(segment, { minLength: 1, maxLength: 3 })
                .map((parts) => parts.join("/")),
              data: fc.oneof(
                fc.uint8Array({ maxLength: 2000 }),
                fc.string({ unit: "grapheme", maxLength: 500 }),
              ),
              compression: fc.constantFrom(
                "store" as const,
                "deflate" as const,
              ),
            }),
            { selector: ({ path }) => path, maxLength: 12 },
          ),
          fc.constantFrom<ZipCodec | undefined>(undefined, webZipCodec),
          async (entries, codec) => {
            const archive = new ZipArchive(
              codec === undefined ? {} : { codec },
            );
            for (const { path, data, compression } of entries) {
              archive.set(path, data, { compression });
            }
            const reread = await ZipArchive.read(await archive.generate());
            expect(reread.paths()).toEqual(entries.map(({ path }) => path));
            for (const { path, data } of entries) {
              const expected =
                typeof data === "string"
                  ? new TextEncoder().encode(data)
                  : data;
              expect(await reread.entry(path)?.bytes()).toEqual(expected);
            }
          },
        ),
        { numRuns: 100, id: "ZIP writer output reads back" },
      );
    },
    propertyTestTimeout(60_000),
  );
});
