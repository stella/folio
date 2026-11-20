import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";

import { ZipArchive } from "@stll/docx-core/zip";
import { FolioDocxReviewer } from "@stll/folio-core/server";

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
import { runPackageRewrite } from "./package-rewrite";
import { sanitizePackage } from "./sanitize";

const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

/** A reused document: comments in the body and a header, revisions, rsids, and properties. */
const REUSED: PackageFixture = {
  body:
    '<w:p w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3" w14:paraId="10000001" w14:textId="77777777">' +
    '<w:commentRangeStart w:id="0"/><w:r w:rsidRPr="00D4E5F6"><w:t>The rate is fixed.</w:t></w:r>' +
    '<w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>' +
    "</w:p>" +
    '<w:p w14:paraId="10000002" w14:textId="77777777"><w:ins w:id="5" w:author="Drafter" w:date="2026-01-01T00:00:00Z">' +
    "<w:r><w:t>Inserted clause.</w:t></w:r></w:ins>" +
    '<w:r><w:t xml:space="preserve">Kept </w:t><w:commentReference w:id="1"/></w:r></w:p>',
  sectPr:
    '<w:sectPr w:rsidR="00A1B2C3"><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>',
  headers: {
    rIdHeader:
      '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Header text</w:t></w:r><w:commentRangeEnd w:id="1"/></w:p>',
  },
  styles:
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rsid w:val="00A1B2C3"/></w:style>',
  settings:
    '<w:zoom w:percent="100"/><w:rsids><w:rsidRoot w:val="00A1B2C3"/><w:rsid w:val="00A1B2C3"/><w:rsid w:val="00D4E5F6"/></w:rsids>',
  comments:
    '<w:comment w:id="0" w:author="Partner A" w:initials="PA"><w:p><w:r><w:t>Check the rate.</w:t></w:r></w:p></w:comment>' +
    '<w:comment w:id="1" w:author="Associate B" w:initials="AB"><w:p><w:r><w:t>Header note.</w:t></w:r></w:p></w:comment>',
  commentsExtended: `<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="20000001" w15:done="0"/></w15:commentsEx>`,
  commentsIds:
    '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w16cid:commentId w16cid:paraId="20000001" w16cid:durableId="1A2B3C4D"/></w16cid:commentsIds>',
  commentsExtensible:
    '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"><w16cex:commentExtensible w16cex:durableId="1A2B3C4D" w16cex:dateUtc="2026-01-01T00:00:00Z"/></w16cex:commentsExtensible>',
  people:
    '<w15:person w15:author="Partner A"><w15:presenceInfo w15:providerId="None" w15:userId="Partner A"/></w15:person>' +
    '<w15:person w15:author="Associate B"><w15:presenceInfo w15:providerId="None" w15:userId="Associate B"/></w15:person>' +
    '<w15:person w15:author="Drafter"><w15:presenceInfo w15:providerId="None" w15:userId="Drafter"/></w15:person>',
  core: "<dc:title>Agreement</dc:title><dc:creator>Partner A</dc:creator><cp:lastModifiedBy>Associate B</cp:lastModifiedBy>",
  app: "<Application>Editor</Application><Company>Example Firm</Company><Manager>Partner A</Manager>",
};

const COMMENT_PARTS = [
  "word/comments.xml",
  "word/commentsExtended.xml",
  "word/commentsIds.xml",
  "word/commentsExtensible.xml",
];

const sanitize = async (
  bytes: Uint8Array,
  args: Record<string, unknown> = {},
) => {
  const pkg = (await openOoxmlPackage(bytes)).unwrap();
  const report = sanitizePackage(pkg, args).unwrap();
  const parts = pkg.changedParts();
  return {
    report,
    parts,
    rewritten: parts.length === 0 ? null : await pkg.toBytes(),
  };
};

const rewrittenTexts = async (
  bytes: Uint8Array | null,
): Promise<ReadonlyMap<string, string>> => {
  if (bytes === null) throw new Error("expected a rewritten package");
  return await packageTexts(bytes);
};

describe("sanitize rewrite", () => {
  test("removes comments, their parts, anchors in every story, and comment-only people", async () => {
    const { report, rewritten } = await sanitize(await buildPackage(REUSED), {
      remove: ["comments"],
    });
    expect(report.comments).toEqual({
      comments: 2,
      anchors: 6,
      authors: ["Associate B", "Partner A"],
      people: ["Associate B", "Partner A"],
    });
    const texts = await rewrittenTexts(rewritten);
    for (const part of COMMENT_PARTS) expect(texts.has(part)).toBe(false);
    const document = texts.get("word/document.xml") ?? "";
    const header = texts.get("word/header-rIdHeader.xml") ?? "";
    for (const story of [document, header]) {
      expect(story).not.toMatch(/comment(RangeStart|RangeEnd|Reference)/u);
    }
    // The run that held only a reference is gone; the one with text is kept.
    expect(document).not.toContain("CommentReference");
    expect(document).toContain("Kept ");
    expect(header).toContain("Header text");
    // The revision author stays in people.xml.
    const people = texts.get("word/people.xml") ?? "";
    expect(people).toContain('w15:author="Drafter"');
    expect(people).not.toContain("Partner A");
    const rels = texts.get("word/_rels/document.xml.rels") ?? "";
    expect(rels).not.toMatch(/comments/iu);
    expect(rels).toContain("people");
    const types = texts.get("[Content_Types].xml") ?? "";
    expect(types).not.toMatch(/comments/iu);
    // rsids were not asked for.
    expect(document).toContain("w:rsidR=");
  });

  test("removes people.xml when no revision names anyone in it", async () => {
    const { rewritten } = await sanitize(
      await buildPackage({
        ...REUSED,
        people:
          '<w15:person w15:author="Partner A"><w15:presenceInfo w15:providerId="None" w15:userId="Partner A"/></w15:person>',
      }),
    );
    const texts = await rewrittenTexts(rewritten);
    expect(texts.has("word/people.xml")).toBe(false);
    expect(texts.get("word/_rels/document.xml.rels")).not.toContain("people");
    expect(texts.get("[Content_Types].xml")).not.toContain("people");
  });

  test("removes rsid attributes, style rsids, and the settings table in any namespace profile", async () => {
    for (const namespace of [undefined, W_STRICT]) {
      const { report, rewritten } = await sanitize(
        await buildPackage({ ...REUSED, ...(namespace && { namespace }) }),
        { remove: ["rsids"] },
      );
      expect(report.rsids).toEqual({ attributes: 4, elements: 2 });
      const texts = await rewrittenTexts(rewritten);
      for (const [part, text] of texts) {
        if (!part.endsWith(".xml")) continue;
        expect(`${part}: ${text}`).not.toMatch(/:rsid/u);
      }
      expect(texts.get("word/settings.xml")).toContain("w:zoom");
      expect(texts.get("word/document.xml")).toContain("The rate is fixed.");
    }
  });

  test("removes identifying properties only when asked", async () => {
    const bytes = await buildPackage(REUSED);
    const defaults = await sanitize(bytes);
    expect(defaults.report.properties).toBeUndefined();
    expect(defaults.parts).not.toContain("docProps/core.xml");

    const { report, rewritten } = await sanitize(bytes, {
      remove: ["properties"],
    });
    expect(report.properties).toEqual([
      { part: "docProps/core.xml", property: "creator", value: "Partner A" },
      {
        part: "docProps/core.xml",
        property: "lastModifiedBy",
        value: "Associate B",
      },
      { part: "docProps/app.xml", property: "company", value: "Example Firm" },
      { part: "docProps/app.xml", property: "manager", value: "Partner A" },
    ]);
    const texts = await rewrittenTexts(rewritten);
    expect(texts.get("docProps/core.xml")).toContain("Agreement");
    expect(texts.get("docProps/core.xml")).not.toContain("Partner A");
    expect(texts.get("docProps/app.xml")).not.toContain("Example Firm");
  });

  test("finds core properties related with the Strict relationship type", async () => {
    const zip = await ZipArchive.read(await buildPackage(REUSED));
    const rels = (await zip.entry("_rels/.rels")?.text()) ?? "";
    zip.set(
      "_rels/.rels",
      rels.replace(
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/metadata/core-properties",
      ),
    );
    const { report } = await sanitize(await zip.generate(), {
      remove: ["properties"],
    });
    expect(report.properties?.map(({ property }) => property)).toEqual([
      "creator",
      "lastModifiedBy",
      "company",
      "manager",
    ]);
  });

  test("is idempotent: a sanitized package has nothing left to remove", async () => {
    const all = { remove: ["comments", "rsids", "properties"] };
    const first = await sanitize(await buildPackage(REUSED), all);
    if (first.rewritten === null) throw new Error("expected a rewrite");
    const second = await sanitize(first.rewritten, all);
    expect(second.parts).toEqual([]);
    expect(second.rewritten).toBeNull();
    // The result still opens, with its text and its revision intact.
    const reviewer = await FolioDocxReviewer.fromBuffer(
      first.rewritten.slice().buffer,
      { session: "legacy", author: "Test" },
    );
    expect(reviewer.getComments()).toEqual([]);
    expect(reviewer.getChanges().length).toBe(1);
    expect(reviewer.getContentAsText()).toContain("The rate is fixed.");
  });

  test("refuses an unknown target", async () => {
    const outcome = await runPackageRewrite(
      "sanitize",
      await buildPackage(REUSED),
      { remove: ["comments", "macros"] },
    );
    expect(outcome.isErr() && outcome.error.code).toBe("invalid_input");
  });
});

describe("folio sanitize", () => {
  let dir = "";
  let cleanup: () => Promise<void> = () => Promise.resolve();

  beforeEach(async () => {
    ({ dir, cleanup } = await makeTempDir());
  });

  afterEach(async () => {
    await cleanup();
  });

  test("reports without a destination, then writes in place, then has nothing left", async () => {
    const file = await writePackage(dir, "reused.docx", REUSED);
    const before = await readFile(file);

    const report = captureIo();
    expect(await runFolioCli(["sanitize", file], report.io)).toBe(0);
    const reported = dataOf(report.stdout());
    expect(reported["result"]).toMatchObject({
      report: { remove: ["comments", "rsids"], comments: { comments: 2 } },
    });
    expect(await readFile(file)).toEqual(before);

    const write = captureIo();
    expect(
      await runFolioCli(
        [
          "sanitize",
          file,
          "--in-place",
          "--expect-version",
          String(reported["fileVersion"]),
        ],
        write.io,
      ),
    ).toBe(0);
    const receipt = dataOf(write.stdout());
    expect(receipt["saveStrategy"]).toBe("part-rewrite");
    expect(receipt["changedParts"]).toContainEqual({
      part: "word/comments.xml",
      change: "removed",
    });

    const again = captureIo();
    expect(
      await runFolioCli(
        [
          "sanitize",
          file,
          "--in-place",
          "--expect-version",
          String(receipt["fileVersion"]),
        ],
        again.io,
      ),
    ).toBe(2);
    expect(envelopeOf(again.stdout())["error"]).toMatchObject({
      code: "operation_rejected",
    });
  });
});
