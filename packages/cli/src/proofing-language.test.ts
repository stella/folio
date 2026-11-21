import {
  afterEach,
  beforeEach,
  describe,
  expect,
  setDefaultTimeout,
  test,
} from "bun:test";
import { readFile } from "node:fs/promises";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import { DISK_SYNC_TEST_TIMEOUT_MS } from "../../../test/disk-sync-timeout";
import { makeTempDir } from "./__tests__/fixtures";
import { captureIo, dataOf, envelopeOf } from "./__tests__/io";
import {
  buildPackage,
  type PackageFixture,
  packageTexts,
  W_STRICT,
  writePackage,
} from "./__tests__/package-fixtures";
import { runFolioCli } from "./cli";
import { openOoxmlPackage } from "./ooxml-package";
import { setProofingLanguage } from "./proofing-language";

setDefaultTimeout(DISK_SYNC_TEST_TIMEOUT_MS);

/** An English template with one German run and a stated paragraph-mark language. */
const ENGLISH: PackageFixture = {
  body:
    '<w:p><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr>' +
    '<w:r><w:rPr><w:b/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr><w:t>Agreement</w:t></w:r>' +
    '<w:r><w:rPr><w:lang w:val="de-DE"/></w:rPr><w:t xml:space="preserve"> Vertrag</w:t></w:r>' +
    "<w:r><w:t> plain</w:t></w:r></w:p>",
  headers: {
    rIdHeader:
      '<w:p><w:r><w:rPr><w:lang w:val="en-GB"/></w:rPr><w:t>Header</w:t></w:r></w:p>',
  },
  sectPr:
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>',
  styles:
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:rPr><w:i/><w:lang w:val="en-US"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/><w:basedOn w:val="Normal"/></w:style>',
  settings:
    '<w:zoom w:percent="100"/><w:themeFontLang w:val="en-US"/><w:compat/>',
};

const run = async (bytes: Uint8Array, args: Record<string, unknown>) => {
  const pkg = (await openOoxmlPackage(bytes)).unwrap();
  const report = setProofingLanguage(pkg, args).unwrap();
  const parts = pkg.changedParts();
  return {
    report,
    parts,
    rewritten: parts.length === 0 ? null : await pkg.toBytes(),
  };
};

const textsOf = async (
  bytes: Uint8Array | null,
): Promise<ReadonlyMap<string, string>> => {
  if (bytes === null) throw new Error("expected a rewritten package");
  return await packageTexts(bytes);
};

describe("proofing language rewrite", () => {
  test("reports the languages each part states, and changes nothing", async () => {
    const { report, parts } = await run(await buildPackage(ENGLISH), {});
    expect(parts).toEqual([]);
    expect(report).toEqual({
      found: [
        {
          part: "word/styles.xml",
          val: { "en-US": 2 },
          eastAsia: { "en-US": 1 },
          bidi: { "ar-SA": 1 },
        },
        { part: "word/settings.xml", val: { "en-US": 1 } },
        {
          part: "word/document.xml",
          val: { "en-US": 2, "de-DE": 1 },
          eastAsia: { "ja-JP": 1 },
        },
        { part: "word/header-rIdHeader.xml", val: { "en-GB": 1 } },
      ],
    });
  });

  test("sets the defaults, the default paragraph style, stated style languages, and settings", async () => {
    const { report, parts, rewritten } = await run(
      await buildPackage(ENGLISH),
      { set: "cs-CZ" },
    );
    expect(report.set).toEqual({
      val: "cs-CZ",
      docDefaults: true,
      styles: 2,
      themeFontLang: true,
      runs: 0,
    });
    expect(parts).toEqual(["word/settings.xml", "word/styles.xml"]);
    const texts = await textsOf(rewritten);
    const styles = texts.get("word/styles.xml") ?? "";
    // Only val changes; eastAsia and bidi keep what they said.
    expect(styles).toContain(
      '<w:lang w:val="cs-CZ" w:eastAsia="en-US" w:bidi="ar-SA"/>',
    );
    // The default paragraph style gains run properties after its qFormat.
    expect(styles).toContain(
      '<w:qFormat/><w:rPr><w:lang w:val="cs-CZ"/></w:rPr></w:style>',
    );
    expect(styles).toContain('<w:i/><w:lang w:val="cs-CZ"/>');
    expect(styles).not.toContain('en-US"/></w:rPr></w:style>');
    expect(texts.get("word/settings.xml")).toContain(
      '<w:themeFontLang w:val="cs-CZ"/>',
    );
  });

  test("with runs, sets every language a run or paragraph mark states", async () => {
    const { report, rewritten } = await run(await buildPackage(ENGLISH), {
      set: "cs-CZ",
      eastAsia: "zh-CN",
      runs: true,
    });
    expect(report.set?.runs).toBe(4);
    const texts = await textsOf(rewritten);
    const document = texts.get("word/document.xml") ?? "";
    expect(document).not.toMatch(/en-US|de-DE|ja-JP/u);
    expect(document).toContain(
      '<w:lang w:val="cs-CZ" w:eastAsia="zh-CN"/></w:rPr><w:t>Agreement',
    );
    expect(texts.get("word/header-rIdHeader.xml")).toContain(
      '<w:lang w:val="cs-CZ" w:eastAsia="zh-CN"/>',
    );
    // The package still opens, with its text.
    if (rewritten === null) throw new Error("expected a rewrite");
    const reviewer = await FolioDocxReviewer.fromBuffer(
      rewritten.slice().buffer,
      { author: "Test" },
    );
    expect(reviewer.getContentAsText()).toContain("Agreement Vertrag plain");
  });

  test("creates the defaults and the settings entry in schema order when absent", async () => {
    const { rewritten } = await run(
      await buildPackage({
        body: "<w:p><w:r><w:t>Text</w:t></w:r></w:p>",
        styles:
          '<w:latentStyles/><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>',
        settings: '<w:zoom w:percent="100"/><w:compat/><w:clrSchemeMapping/>',
        namespace: W_STRICT,
      }),
      { set: "sk-SK" },
    );
    const texts = await textsOf(rewritten);
    expect(texts.get("word/styles.xml")).toContain(
      '<w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="sk-SK"/></w:rPr></w:rPrDefault></w:docDefaults><w:latentStyles/>',
    );
    expect(texts.get("word/settings.xml")).toContain(
      '<w:zoom w:percent="100"/><w:compat/><w:themeFontLang w:val="sk-SK"/><w:clrSchemeMapping/>',
    );
  });

  test("is idempotent: setting the same language again changes nothing", async () => {
    const args = { set: "cs-CZ", runs: true };
    const first = await run(await buildPackage(ENGLISH), args);
    if (first.rewritten === null) throw new Error("expected a rewrite");
    const second = await run(first.rewritten, args);
    expect(second.parts).toEqual([]);
    expect(second.report.set).toMatchObject({
      docDefaults: false,
      styles: 0,
      themeFontLang: false,
      runs: 0,
    });
  });

  test("refuses a value that is not a language tag", async () => {
    const pkg = (await openOoxmlPackage(await buildPackage(ENGLISH))).unwrap();
    for (const args of [{ set: "cs CZ" }, { bidi: "" }, { runs: true }]) {
      const outcome = setProofingLanguage(pkg, args);
      expect(outcome.isErr() && outcome.error.code).toBe("invalid_input");
    }
  });
});

describe("folio lang", () => {
  let dir = "";
  let cleanup: () => Promise<void> = () => Promise.resolve();

  beforeEach(async () => {
    ({ dir, cleanup } = await makeTempDir());
  });

  afterEach(async () => {
    await cleanup();
  });

  test("reports, then sets the language in place", async () => {
    const file = await writePackage(dir, "template.docx", ENGLISH);
    const before = await readFile(file);

    const report = captureIo();
    expect(await runFolioCli(["lang", file, "--set", "cs-CZ"], report.io)).toBe(
      0,
    );
    const reported = dataOf(report.stdout());
    expect(reported["result"]).toMatchObject({
      parts: ["word/settings.xml", "word/styles.xml"],
    });
    expect(await readFile(file)).toEqual(before);

    const write = captureIo();
    expect(
      await runFolioCli(
        [
          "lang",
          file,
          "--set",
          "cs-CZ",
          "--runs",
          "--in-place",
          "--expect-version",
          String(reported["fileVersion"]),
        ],
        write.io,
      ),
    ).toBe(0);
    expect(dataOf(write.stdout())["saveStrategy"]).toBe("part-rewrite");
    const texts = await packageTexts(await readFile(file));
    expect(texts.get("word/document.xml")).not.toContain("de-DE");

    const bad = captureIo();
    expect(await runFolioCli(["lang", file, "--set", "??"], bad.io)).toBe(2);
    expect(envelopeOf(bad.stdout())["error"]).toMatchObject({
      code: "invalid_input",
    });
  });
});
