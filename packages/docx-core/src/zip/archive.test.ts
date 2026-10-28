import { describe, expect, test } from "bun:test";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { join, resolve } from "node:path";

import { readZipRecords, systemZipCheck } from "./__tests__/support";
import { ZipArchive, ZipError } from "./archive";
import { webZipCodec } from "./codec";

const REPOSITORY_ROOT = resolve(import.meta.dir, "../../../..");

/**
 * Entry hashes JSZip produced for every office fixture in the repository,
 * recorded before the archive module replaced it.
 */
const JSZIP_ENTRIES: Record<string, Record<string, string>> = JSON.parse(
  readFileSync(
    join(import.meta.dir, "__fixtures__/jszip-entries.json"),
    "utf8",
  ),
);

const shortHash = (bytes: Uint8Array): string =>
  createHash("sha256").update(bytes).digest("hex").slice(0, 16);

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(REPOSITORY_ROOT, name)));

const entryHashes = async (
  archive: ZipArchive,
): Promise<Record<string, string>> => {
  const hashes: Record<string, string> = {};
  for (const entry of archive.entries()) {
    hashes[entry.path] = shortHash(await entry.bytes());
  }
  return hashes;
};

describe("fixture archives read the entries JSZip read", () => {
  const fixtures = Object.keys(JSZIP_ENTRIES);

  test("the recorded set covers the repository's office fixtures", () => {
    expect(fixtures.length).toBeGreaterThan(70);
  });

  for (const codec of ["platform", "web"] as const) {
    test(`${codec} codec`, async () => {
      for (const name of fixtures) {
        const archive = await ZipArchive.read(fixture(name), {
          ...(codec === "web" ? { codec: webZipCodec } : {}),
        });
        expect({ name, entries: await entryHashes(archive) }).toEqual({
          name,
          entries: JSZIP_ENTRIES[name] ?? {},
        });
        // Entry order is the central directory's, as JSZip reported it.
        expect(archive.paths()).toEqual(Object.keys(JSZIP_ENTRIES[name] ?? {}));
      }
    });
  }
});

describe("writing", () => {
  const SAMPLE = "parity/fixtures/layout-kitchen-sink.docx";

  test("unchanged entries keep their compressed bytes and metadata", async () => {
    const source = fixture(SAMPLE);
    const archive = await ZipArchive.read(source);
    archive.set("word/document.xml", "<w:document/>");
    const output = await archive.generate();

    const before = readZipRecords(source);
    const after = readZipRecords(output);
    expect(after.map(({ path }) => path)).toEqual(
      before.map(({ path }) => path),
    );
    for (const record of after) {
      // SAFETY: the path lists are equal, so every output record has a source.
      const original = before.find(({ path }) => path === record.path)!;
      if (record.path === "word/document.xml") {
        expect(record.data).not.toEqual(original.data);
        continue;
      }
      expect(record.data).toEqual(original.data);
      expect(record.method).toBe(original.method);
      expect(record.crc32).toBe(original.crc32);
      expect(record.dosDateTime).toEqual(original.dosDateTime);
    }
    const reread = await ZipArchive.read(output);
    expect(await reread.entry("word/document.xml")?.text()).toBe(
      "<w:document/>",
    );
  });

  test("a replaced entry keeps its position and no directory entries appear", async () => {
    const archive = new ZipArchive();
    archive.set("word/media/nested/image.png", Uint8Array.of(1, 2, 3));
    archive.set("[Content_Types].xml", "<Types/>");
    archive.set("word/media/nested/image.png", Uint8Array.of(4));
    const reread = await ZipArchive.read(await archive.generate());
    expect(reread.paths()).toEqual([
      "word/media/nested/image.png",
      "[Content_Types].xml",
    ]);
    expect(await reread.entry("word/media/nested/image.png")?.bytes()).toEqual(
      Uint8Array.of(4),
    );
  });

  test("output is reproducible: new entries carry a fixed timestamp", async () => {
    const build = async () => {
      const archive = new ZipArchive();
      archive.set("docProps/core.xml", "<cp:coreProperties/>");
      archive.set("word/document.xml", "<w:document/>", {
        compression: "store",
      });
      return await archive.generate();
    };
    const first = await build();
    expect(await build()).toEqual(first);
    for (const record of readZipRecords(first)) {
      // 2000-01-01 00:00:00 in DOS date and time fields.
      expect(record.dosDateTime).toEqual({
        date: (20 << 9) | (1 << 5) | 1,
        time: 0,
      });
    }
  });

  test("an explicit timestamp is written in UTC", async () => {
    const archive = new ZipArchive();
    archive.set("a.txt", "a", {
      modified: new Date(Date.UTC(2021, 3, 5, 6, 7, 8)),
    });
    const output = await archive.generate();
    const [record] = readZipRecords(output);
    expect(record?.dosDateTime).toEqual({
      date: (41 << 9) | (4 << 5) | 5,
      time: (6 << 11) | (7 << 5) | 4,
    });
    // Read back at DOS resolution: even seconds survive exactly.
    expect((await ZipArchive.read(output)).entry("a.txt")?.modified).toEqual(
      new Date(Date.UTC(2021, 3, 5, 6, 7, 8)),
    );
  });

  test("a generate timestamp replaces every entry's, read or set", async () => {
    const source = await ZipArchive.read(fixture(SAMPLE));
    source.set("word/document.xml", "<w:document/>", {
      modified: new Date(Date.UTC(2021, 0, 1)),
    });
    const output = await source.generate({
      modified: new Date(Date.UTC(2024, 1, 2)),
    });
    const stamp = { date: (44 << 9) | (2 << 5) | 2, time: 0 };
    for (const record of readZipRecords(output)) {
      expect(record.dosDateTime).toEqual(stamp);
    }
  });

  test("a clone edits independently and shares the read bytes", async () => {
    const original = await ZipArchive.read(fixture(SAMPLE));
    const copy = original.clone();
    copy.set("word/document.xml", "<w:document/>");
    copy.delete("[Content_Types].xml");
    expect(original.has("[Content_Types].xml")).toBe(true);
    expect(await original.entry("word/document.xml")?.text()).not.toBe(
      "<w:document/>",
    );
    const passedThrough = readZipRecords(await copy.generate()).find(
      ({ path }) => path === "word/styles.xml",
    );
    const sourceRecord = readZipRecords(fixture(SAMPLE)).find(
      ({ path }) => path === "word/styles.xml",
    );
    expect(passedThrough?.data).toEqual(sourceRecord?.data);
  });

  test("an entry's deflate level overrides the archive's", async () => {
    const words = ["smlouva", "strana", "článek", "lhůta", "plnění", "cena"];
    let seed = 7;
    const text = Array.from({ length: 50_000 }, () => {
      seed = (seed * 1_103_515_245 + 12_345) % 2_147_483_648;
      return words[seed % words.length];
    }).join(" ");
    const archive = new ZipArchive();
    archive.set("fast.txt", text, { level: 1 });
    archive.set("default.txt", text);
    const [fast, normal] = readZipRecords(await archive.generate({ level: 9 }));
    expect(fast?.data.byteLength).toBeGreaterThan(normal?.data.byteLength ?? 0);
  });

  test("compression applies per entry, archive default first", async () => {
    const text = "repetitive ".repeat(1000);
    const archive = new ZipArchive();
    archive.set("deflated.txt", text);
    archive.set("stored.txt", text, { compression: "store" });
    const deflated = readZipRecords(await archive.generate());
    expect(deflated.map(({ method }) => method)).toEqual([8, 0]);
    const stored = readZipRecords(
      await archive.generate({ compression: "store" }),
    );
    expect(stored.map(({ method }) => method)).toEqual([0, 0]);
  });

  test("text is written as UTF-8 in one piece", async () => {
    const text = `${"a".repeat(16 * 1024 - 1)}\u{1F600}end`;
    const archive = new ZipArchive();
    archive.set("word/document.xml", text);
    const reread = await ZipArchive.read(await archive.generate());
    expect(await reread.entry("word/document.xml")?.text()).toBe(text);
  });

  test("a byte-order mark survives a text read", async () => {
    const archive = new ZipArchive();
    archive.set("bom.xml", Uint8Array.of(0xef, 0xbb, 0xbf, 0x3c));
    expect(await archive.entry("bom.xml")?.text()).toBe("\uFEFF<");
  });

  test("written data is copied, so later mutation does not leak in", async () => {
    const data = Uint8Array.of(1, 2, 3);
    const archive = new ZipArchive();
    archive.set("a.bin", data);
    data[0] = 9;
    expect(await archive.entry("a.bin")?.bytes()).toEqual(
      Uint8Array.of(1, 2, 3),
    );
  });

  test("entry names must be relative file paths", () => {
    const archive = new ZipArchive();
    for (const name of [
      "",
      "/abs.xml",
      "../up.xml",
      "a/../../b",
      "a\\b",
      "C:/x",
      "nul\0.xml",
      "folder/",
    ]) {
      expect(() => archive.set(name, "x")).toThrow(ZipError);
    }
  });

  test("non-ASCII names are flagged UTF-8 and read back", async () => {
    const archive = new ZipArchive();
    archive.set("word/media/obrázek.png", Uint8Array.of(1));
    const output = await archive.generate();
    const [record] = readZipRecords(output);
    expect((record?.flags ?? 0) & 0x08_00).toBe(0x08_00);
    expect((await ZipArchive.read(output)).paths()).toEqual([
      "word/media/obrázek.png",
    ]);
  });

  test("the web codec writes archives the platform reads", async () => {
    const archive = new ZipArchive({ codec: webZipCodec });
    archive.set("word/document.xml", "<w:document>".repeat(500));
    const reread = await ZipArchive.read(await archive.generate());
    expect(await reread.entry("word/document.xml")?.text()).toBe(
      "<w:document>".repeat(500),
    );
  });

  test.if(systemZipCheck.available)(
    "Info-ZIP and Python accept the output",
    async () => {
      const archive = await ZipArchive.read(fixture(SAMPLE));
      archive.set("word/document.xml", "<w:document/>");
      archive.set("word/media/obrázek.png", Uint8Array.of(1, 2, 3), {
        compression: "store",
      });
      expect(systemZipCheck.check(await archive.generate())).toEqual({
        ok: true,
      });
    },
  );
});
