import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import { ZipArchive } from "@stll/docx-core/zip";
import { createZipStream } from "@stll/docx-core/zip/stream";

import { withFixedPackageDates } from "../../compare/reproducible-package";
import { ensureParaIds } from "../ensureParaIds";
import {
  FOLIO_DOCUMENT_PRIVACY_TRANSFORMS,
  rewriteDocxMetadataPrivacy,
} from "../metadataPrivacy";
import { parseDocx } from "../parser";
import { repackDocx } from "../rezip";
import { attemptSelectiveSave } from "../selectiveSave";

const FIXTURES_DIR = path.join(import.meta.dir, "__fixtures__", "corpus");
const FIXTURE_FILES = readdirSync(FIXTURES_DIR)
  .filter((name) => name.endsWith(".docx"))
  .sort();
const CORE_PROPERTIES_PATH = "docProps/core.xml";
const PACKAGE_DATE = new Date("2000-01-01T00:00:00.000Z");
const CORE_PROPERTIES_XML = `<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Fixture title</dc:title><dc:subject>Fixture subject</dc:subject><dc:creator>Fixture creator</dc:creator><cp:keywords>fixture</cp:keywords><dc:description>Fixture description</dc:description><cp:lastModifiedBy>Fixture modifier</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">1999-01-01T00:00:00.000Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">1999-01-01T00:00:00.000Z</dcterms:modified></cp:coreProperties>`;

/**
 * Every name in the central directory, directories included: the archive
 * reader skips directory records, so the oracle walks the records itself
 * rather than hide the regression it exists to catch.
 */
const entryNames = async (
  buffer: ArrayBuffer | Uint8Array,
): Promise<string[]> => {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = bytes.byteLength - 22;
  while (end >= 0 && view.getUint32(end, true) !== 0x06_05_4b_50) end--;
  const names: string[] = [];
  let cursor = view.getUint32(end + 16, true);
  for (let index = 0; index < view.getUint16(end + 10, true); index++) {
    const nameLength = view.getUint16(cursor + 28, true);
    names.push(
      new TextDecoder().decode(
        bytes.subarray(cursor + 46, cursor + 46 + nameLength),
      ),
    );
    cursor +=
      46 +
      nameLength +
      view.getUint16(cursor + 30, true) +
      view.getUint16(cursor + 32, true);
  }
  return names.toSorted();
};

type EntryNamesPreservedOptions = {
  source: ArrayBuffer;
  saved: ArrayBuffer | Uint8Array;
};

const expectEntryNamesPreserved = async ({
  source,
  saved,
}: EntryNamesPreservedOptions): Promise<void> => {
  expect(await entryNames(saved)).toEqual(await entryNames(source));
};

const directoryFreeFixture = async (filename: string): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(
    readFileSync(path.join(FIXTURES_DIR, filename)),
  );
  return zip.generate().then((bytes) => bytes.buffer);
};

const withPrivateMetadata = async (
  buffer: ArrayBuffer,
): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(buffer);
  // The corpus currently has no core properties. Seed them in the source so
  // privacy, date restamping, and selective saves must actually rewrite a part.
  zip.set(CORE_PROPERTIES_PATH, CORE_PROPERTIES_XML, {
    modified: PACKAGE_DATE,
  });
  return zip.generate().then((bytes) => bytes.buffer);
};

const corePropertiesText = async (buffer: ArrayBuffer): Promise<string> => {
  const part = (await ZipArchive.read(buffer)).entry(CORE_PROPERTIES_PATH);
  if (part === undefined)
    panic("Expected source core properties to survive save");
  return part.text();
};

describe("save paths preserve corpus package entry names", () => {
  test("the corpus is present and the oracle detects invented directories", async () => {
    expect(FIXTURE_FILES.length).toBeGreaterThan(0);
    const fixture = FIXTURE_FILES.at(0);
    if (fixture === undefined) panic("Expected a corpus fixture");
    const source = await directoryFreeFixture(fixture);
    expect((await entryNames(source)).some((name) => name.endsWith("/"))).toBe(
      false,
    );
    const zip = await ZipArchive.read(source);
    // Deliberately invent a parent directory entry to prove the oracle sees
    // it, even though every file part remains unchanged.
    const parts = await Promise.all(
      zip.entries().map(async (entry) => ({
        path: entry.path,
        data: await entry.bytes(),
      })),
    );
    const wrong = new Uint8Array(
      await new Response(
        createZipStream([{ path: "word/" }, ...parts]),
      ).arrayBuffer(),
    );
    expect(await entryNames(wrong)).toContain("word/");
    expect(await entryNames(wrong)).not.toEqual(await entryNames(source));
  });

  test.each(FIXTURE_FILES)(
    "full repack and paragraph-ID writes preserve entries (%s)",
    async (filename) => {
      const source = await directoryFreeFixture(filename);
      expect(
        (await entryNames(source)).some((name) => name.endsWith("/")),
      ).toBe(false);
      await expectEntryNamesPreserved({
        source,
        saved: (await ensureParaIds(source)).docx,
      });
      const document = await parseDocx(source, { preloadFonts: false });
      await expectEntryNamesPreserved({
        source,
        saved: await repackDocx(document, { updateModifiedDate: false }),
      });
    },
  );

  test.each(FIXTURE_FILES)(
    "selective save, fixed dates, and every privacy transform preserve entries (%s)",
    async (filename) => {
      const source = await withPrivateMetadata(
        await directoryFreeFixture(filename),
      );
      const document = await parseDocx(source, { preloadFonts: false });
      const repacked = await repackDocx(document);
      await expectEntryNamesPreserved({ source, saved: repacked });
      expect(await corePropertiesText(repacked)).not.toBe(CORE_PROPERTIES_XML);
      const selective = await attemptSelectiveSave(document, source, {
        changedParaIds: new Set(),
        structuralChange: false,
        hasUntrackedChanges: false,
      });
      // These no-op models need no resource materialization. A refusal must fail
      // the case, rather than silently dropping that fixture's save coverage.
      expect(selective).not.toBeNull();
      if (selective === null) panic(`Selective save refused ${filename}`);
      await expectEntryNamesPreserved({ source, saved: selective });
      expect(await corePropertiesText(selective)).not.toBe(CORE_PROPERTIES_XML);

      const fixedDates = await withFixedPackageDates(source, PACKAGE_DATE);
      await expectEntryNamesPreserved({ source, saved: fixedDates });
      expect(await corePropertiesText(fixedDates)).toContain(
        PACKAGE_DATE.toISOString(),
      );

      for (const transforms of [
        ...FOLIO_DOCUMENT_PRIVACY_TRANSFORMS.map((transform) => [transform]),
        FOLIO_DOCUMENT_PRIVACY_TRANSFORMS,
      ]) {
        // oxlint-disable-next-line no-await-in-loop -- exercise every privacy transform against each corpus fixture without concurrent archive inflation
        const rewritten = await rewriteDocxMetadataPrivacy(source, {
          transforms,
        });
        expect(
          rewritten.privacyReport.removedMetadataProperties.length,
        ).toBeGreaterThan(0);
        // oxlint-disable-next-line no-await-in-loop -- the equality oracle includes all directory entries for this transform's completed save
        await expectEntryNamesPreserved({ source, saved: rewritten.buffer });
      }
    },
  );
});
