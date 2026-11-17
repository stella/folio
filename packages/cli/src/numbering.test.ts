import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import { makeTempDir } from "./__tests__/fixtures";
import { captureIo, dataOf } from "./__tests__/io";
import {
  buildPackage,
  type PackageFixture,
  packageTexts,
} from "./__tests__/package-fixtures";
import { runFolioCli } from "./cli";
import { runPackageRewrite } from "./package-rewrite";

const numbered = (id: string, text: string, numPr = "", style = ""): string =>
  `<w:p w14:paraId="${id}" w14:textId="77777777"><w:pPr>${style === "" ? "" : `<w:pStyle w:val="${style}"/>`}${numPr}</w:pPr>` +
  `<w:r><w:t>${text}</w:t></w:r></w:p>`;
const numPr = (numId: number, level: number): string =>
  `<w:numPr><w:ilvl w:val="${String(level)}"/><w:numId w:val="${String(numId)}"/></w:numPr>`;

const LEVELS =
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
  '<w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
  "<w:rPr><w:b/></w:rPr></w:lvl>" +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="space"/><w:lvlText w:val="%1.%2"/>' +
  '<w:pPr><w:ind w:left="1440" w:hanging="720"/></w:pPr></w:lvl>';

/** Numbered clauses across two sections, one restarting list, a style-numbered paragraph, and bullets. */
const FILING: PackageFixture = {
  body:
    numbered("10000001", "Alpha", numPr(1, 0)) +
    numbered("10000002", "Beta", numPr(1, 1)) +
    numbered("10000003", "Gamma", numPr(1, 0)) +
    '<w:p w14:paraId="10000004" w14:textId="77777777"><w:pPr><w:sectPr/></w:pPr><w:r><w:t>Section break</w:t></w:r></w:p>' +
    numbered("10000005", "Delta", numPr(2, 0)) +
    numbered("10000006", "Epsilon", "", "ListClause") +
    numbered("10000007", "Bullet", numPr(4, 0)) +
    numbered("10000008", "[9.]\tLiteral"),
  numbering:
    `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="multilevel"/>${LEVELS}</w:abstractNum>` +
    '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>' +
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>' +
    '<w:num w:numId="4"><w:abstractNumId w:val="2"/></w:num>',
  styles:
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    `<w:style w:type="paragraph" w:styleId="ListClause"><w:name w:val="List Clause"/><w:basedOn w:val="Normal"/><w:pPr>${numPr(1, 0)}</w:pPr></w:style>`,
};

const rewrite = async (bytes: Uint8Array, args: Record<string, unknown>) => {
  const outcome = (await runPackageRewrite("numbering", bytes, args)).unwrap();
  return { report: outcome.report, bytes: await outcome.bytes() };
};

const documentOf = async (bytes: Uint8Array | null): Promise<string> => {
  if (bytes === null) throw new Error("expected a rewritten package");
  return (await packageTexts(bytes)).get("word/document.xml") ?? "";
};

const paragraphXml = (document: string, id: string): string =>
  /<w:p [^>]*>(?:(?!<\/w:p>).)*<\/w:p>/gsu
    .exec(document.slice(document.indexOf(`w14:paraId="${id}"`) - 6))
    ?.at(0) ?? "";

const blocksOf = async (bytes: Uint8Array | null) => {
  if (bytes === null) throw new Error("expected a rewritten package");
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer);
  return reviewer.snapshot().blocks;
};

describe("literal numbering", () => {
  test("writes the displayed labels and removes list membership, leaving bullets", async () => {
    const { report, bytes } = await rewrite(await buildPackage(FILING), {});
    expect(report).toEqual({
      action: "literal",
      paragraphs: [
        { blockId: "10000001", numId: 1, level: 0, label: "1." },
        { blockId: "10000002", numId: 1, level: 1, label: "1.1" },
        { blockId: "10000003", numId: 1, level: 0, label: "2." },
        { blockId: "10000005", numId: 2, level: 0, label: "1." },
        // Instances of one abstract definition share its count: after the
        // restart, numId 1 continues from it.
        { blockId: "10000006", numId: 1, level: 0, label: "2." },
      ],
    });
    const blocks = await blocksOf(bytes);
    const texts = blocks.map(({ text }) => text);
    expect(texts.slice(0, 3)).toEqual(["1.\tAlpha", "1.1 Beta", "2.\tGamma"]);
    expect(texts[5]).toBe("2.\tEpsilon");
    // Nothing the rewrite touched is still in a list; the bullet still is.
    for (const block of blocks) {
      if (block.kind === "diagnostic") continue;
      expect(block.listReference !== undefined).toBe(block.id === "10000007");
    }
  });

  test("keeps the level's indentation, tab stops, and label formatting as direct formatting", async () => {
    const { bytes } = await rewrite(await buildPackage(FILING), {
      blockIds: ["10000001", "10000006"],
    });
    const document = await documentOf(bytes);
    const first = paragraphXml(document, "10000001");
    expect(first).not.toContain("<w:numPr>");
    expect(first).toContain(
      '<w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs>',
    );
    expect(first).toContain('<w:ind w:left="720" w:hanging="360"/>');
    expect(first).toContain(
      '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">1.</w:t><w:tab/></w:r>',
    );
    // Numbering a style supplies is switched off rather than left to show through.
    expect(paragraphXml(document, "10000006")).toContain(
      '<w:numPr><w:numId w:val="0"/></w:numPr>',
    );
    // Paragraphs not asked for keep their numbering.
    expect(paragraphXml(document, "10000003")).toContain("<w:numPr>");
  });

  test("continuous house-style labels ignore restarts", async () => {
    const { report } = await rewrite(await buildPackage(FILING), {
      labels: "continuous",
      format: "[{n}.]",
      numIds: [1, 2],
    });
    expect(report).toMatchObject({
      paragraphs: ["[1.]", "[1.]", "[2.]", "[3.]", "[4.]"].map((label) => ({
        label,
      })),
    });
  });
});

describe("apply numbering", () => {
  test("puts paragraphs into a list and strips the literal label they start with", async () => {
    const { report, bytes } = await rewrite(await buildPackage(FILING), {
      action: "apply",
      numId: 1,
      level: 0,
      blockIds: ["10000008", "10000004"],
      strip: String.raw`\[\d+\.\]\t`,
    });
    expect(report).toEqual({
      action: "apply",
      paragraphs: [
        { blockId: "10000008", numId: 1, level: 0, label: "[9.]\t" },
        { blockId: "10000004", numId: 1, level: 0 },
      ],
    });
    const blocks = await blocksOf(bytes);
    const literal = blocks.find(({ id }) => id === "10000008");
    expect(literal?.text).toBe("Literal");
    expect(literal?.kind !== "diagnostic" && literal?.listReference).toEqual({
      numId: 1,
      level: 0,
    });
  });

  test("literal then apply round-trips list membership", async () => {
    const literal = await rewrite(await buildPackage(FILING), {
      blockIds: ["10000001", "10000003"],
    });
    if (literal.bytes === null) throw new Error("expected a rewrite");
    const applied = await rewrite(literal.bytes, {
      action: "apply",
      numId: 1,
      blockIds: ["10000001", "10000003"],
      strip: String.raw`\d+\.\t`,
    });
    const blocks = await blocksOf(applied.bytes);
    const original = await blocksOf(await buildPackage(FILING));
    for (const id of ["10000001", "10000003"]) {
      const after = blocks.find((block) => block.id === id);
      const before = original.find((block) => block.id === id);
      expect(after?.text).toBe(before?.text);
      expect(after?.displayLabel).toBe(before?.displayLabel);
    }
  });

  test("refuses an undefined list, an unknown paragraph, and a missing target", async () => {
    const bytes = await buildPackage(FILING);
    for (const [args, code] of [
      [{ action: "apply", numId: 9, blockIds: ["10000001"] }, "not_found"],
      [{ action: "apply", numId: 1, blockIds: ["DEADBEEF"] }, "not_found"],
      [{ action: "apply", numId: 1 }, "invalid_input"],
      [{ blockIds: ["10000007"] }, "not_found"],
      [{ format: "no counter" }, "invalid_input"],
    ] as const) {
      const outcome = await runPackageRewrite("numbering", bytes, args);
      expect(outcome.isErr() && outcome.error.code).toBe(code);
    }
  });
});

describe("folio literal-numbering", () => {
  let dir = "";
  let cleanup: () => Promise<void> = () => Promise.resolve();

  beforeEach(async () => {
    ({ dir, cleanup } = await makeTempDir());
  });

  afterEach(async () => {
    await cleanup();
  });

  test("reports, then writes the labels to a new file", async () => {
    const file = path.join(dir, "filing.docx");
    await writeFile(file, await buildPackage(FILING));
    const report = captureIo();
    expect(
      await runFolioCli(
        [
          "literal-numbering",
          file,
          "--labels",
          "continuous",
          "--format",
          "[{n}.]",
        ],
        report.io,
      ),
    ).toBe(0);
    const reported = dataOf(report.stdout());
    const out = path.join(dir, "literal.docx");
    const write = captureIo();
    expect(
      await runFolioCli(
        [
          "literal-numbering",
          file,
          "--labels",
          "continuous",
          "--format",
          "[{n}.]",
          "-o",
          out,
          "--expect-version",
          String(reported["fileVersion"]),
        ],
        write.io,
      ),
    ).toBe(0);
    expect(dataOf(write.stdout())["saveStrategy"]).toBe("part-rewrite");
    const blocks = await blocksOf(new Uint8Array(await readFile(out)));
    expect(blocks.at(0)?.text).toBe("[1.]\tAlpha");
  });
});
