import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import { makeTempDir } from "./__tests__/fixtures";
import { captureIo, dataOf, envelopeOf } from "./__tests__/io";
import {
  buildPackage,
  type PackageFixture,
  packageTexts,
} from "./__tests__/package-fixtures";
import { runFolioCli } from "./cli";
import { EXIT_CODES } from "./errors";
import { fillPlaceholders, inspectPlaceholders } from "./placeholders";

const paragraph = (id: string, runs: string): string =>
  `<w:p w14:paraId="${id}" w14:textId="77777777">${runs}</w:p>`;
const run = (text: string, bold = false): string =>
  `<w:r>${bold ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A draft with placeholders in the body, a table, a header, and a footer. */
const DRAFT: PackageFixture = {
  body:
    paragraph(
      "10000001",
      run("Between ") +
        run("{{cli", true) +
        run("ent}}", true) +
        run(" and [•]."),
    ) +
    paragraph("10000002", run("Signed on [[signing date]]. TODO check.")) +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tr><w:tc>' +
    paragraph("10000003", run("Fee: [•]")) +
    "</w:tc></w:tr></w:tbl>" +
    paragraph("10000004", run("No placeholders here; [2026] is a year.")),
  headers: { rIdHeader: paragraph("20000001", run("Ref {{matter}}")) },
  footers: { rIdFooter: paragraph("30000001", run("Draft […]")) },
  sectPr:
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr>',
};

const reviewerOf = async (bytes: Uint8Array): Promise<FolioDocxReviewer> =>
  await FolioDocxReviewer.fromBuffer(bytes.slice().buffer, {
    author: "Drafter",
  });

describe("placeholders", () => {
  test("lists every placeholder in every story with its block, name, and context", async () => {
    const report = inspectPlaceholders(
      await reviewerOf(await buildPackage(DRAFT)),
      {},
    ).unwrap();
    expect(
      report.placeholders.map(({ story, blockId, text, name }) => ({
        story: story.type,
        blockId,
        text,
        ...(name !== undefined && { name }),
      })),
    ).toEqual([
      {
        story: "main",
        blockId: "10000001",
        text: "{{client}}",
        name: "client",
      },
      { story: "main", blockId: "10000001", text: "[•]" },
      {
        story: "main",
        blockId: "10000002",
        text: "[[signing date]]",
        name: "signing date",
      },
      { story: "main", blockId: "10000002", text: "TODO" },
      { story: "main", blockId: "10000003", text: "[•]" },
      {
        story: "header",
        blockId: "20000001",
        text: "{{matter}}",
        name: "matter",
      },
      { story: "footer", blockId: "30000001", text: "[…]" },
    ]);
    expect(report.placeholders[0]?.context).toBe("Between {{client}} and [•].");
    expect(report.unresolved).toBe(7);
    expect(report.filled).toBe(0);
  });

  test("shows the value that would fill each, by name or by exact text", async () => {
    const report = inspectPlaceholders(
      await reviewerOf(await buildPackage(DRAFT)),
      { fill: { client: "Acme Ltd", "[•]": "Beta plc", matter: "M-17" } },
    ).unwrap();
    expect(
      report.placeholders.flatMap(({ text, fill }) =>
        fill === undefined ? [] : [[text, fill]],
      ),
    ).toEqual([
      ["{{client}}", "Acme Ltd"],
      ["[•]", "Beta plc"],
      ["[•]", "Beta plc"],
      ["{{matter}}", "M-17"],
    ]);
    expect(report.unresolved).toBe(3);
  });

  test("custom patterns replace the defaults", async () => {
    const report = inspectPlaceholders(
      await reviewerOf(await buildPackage(DRAFT)),
      { patterns: [String.raw`\[(?<name>\d{4})\]`] },
    ).unwrap();
    expect(report.placeholders.map(({ text, name }) => [text, name])).toEqual([
      ["[2026]", "2026"],
    ]);
  });

  test("refuses invalid patterns and values that are placeholders themselves", async () => {
    const reviewer = await reviewerOf(await buildPackage(DRAFT));
    for (const args of [
      { patterns: ["("] },
      { patterns: ["x*"] },
      { fill: { client: "{{client}}" } },
      { fill: { client: 7 } },
    ]) {
      const outcome = inspectPlaceholders(reviewer, args);
      expect(outcome.isErr() && outcome.error.code).toBe("invalid_input");
    }
  });

  test("fills as tracked changes across stories, keeps run formatting, and counts what is left", async () => {
    const reviewer = await reviewerOf(await buildPackage(DRAFT));
    const report = fillPlaceholders(
      reviewer,
      { fill: { client: "Acme Ltd", "[•]": "Beta plc", matter: "M-17" } },
      { mode: "tracked-changes", date: "2026-01-02T03:04:05Z", idSeed: 100 },
    ).unwrap();
    expect(report.filled).toBe(4);
    expect(report.unresolved).toBe(3);
    expect(report.placeholders.map(({ text }) => text)).toEqual([
      "[[signing date]]",
      "TODO",
      "[…]",
    ]);
    const texts = await packageTexts(new Uint8Array(await reviewer.toBuffer()));
    const document = texts.get("word/document.xml") ?? "";
    expect(document).toContain("<w:ins");
    expect(document).toContain("<w:delText");
    expect(document).toContain("Acme Ltd");
    expect(texts.get("word/header-rIdHeader.xml")).toContain("M-17");
    // Accepting everything leaves the filled text, with the bold run still bold.
    const accepted = await reviewerOf(
      new Uint8Array(await reviewer.toBuffer()),
    );
    accepted.acceptAll();
    expect(accepted.getContentAsText()).toContain(
      "Between Acme Ltd and Beta plc.",
    );
    const acceptedDocument =
      (await packageTexts(new Uint8Array(await accepted.toBuffer()))).get(
        "word/document.xml",
      ) ?? "";
    expect(acceptedDocument).toMatch(/<w:b\/>.*Acme Ltd/su);
  });

  test("fills directly when asked", async () => {
    const reviewer = await reviewerOf(await buildPackage(DRAFT));
    fillPlaceholders(
      reviewer,
      { fill: { "[•]": "Beta plc" } },
      { mode: "direct", date: "2026-01-02T03:04:05Z", idSeed: 100 },
    ).unwrap();
    const document =
      (await packageTexts(new Uint8Array(await reviewer.toBuffer()))).get(
        "word/document.xml",
      ) ?? "";
    expect(document).not.toContain("<w:ins");
    expect(document).not.toContain("[•]");
    expect(reviewer.getContentAsText()).toContain("Fee: Beta plc");
  });
});

describe("folio placeholders", () => {
  let dir = "";
  let cleanup: () => Promise<void> = () => Promise.resolve();

  beforeEach(async () => {
    ({ dir, cleanup } = await makeTempDir());
  });

  afterEach(async () => {
    await cleanup();
  });

  test("exits non-zero while placeholders remain, and zero once a fill resolves them", async () => {
    const file = path.join(dir, "draft.docx");
    await writeFile(file, await buildPackage(DRAFT));

    const list = captureIo();
    expect(await runFolioCli(["placeholders", file], list.io)).toBe(
      EXIT_CODES.unresolved,
    );
    const listed = dataOf(list.stdout());
    expect(listed["result"]).toMatchObject({ unresolved: 7 });

    const fields = path.join(dir, "fields.json");
    await writeFile(
      fields,
      JSON.stringify({
        client: "Acme Ltd",
        "[•]": "Beta plc",
        "signing date": "1 March 2027",
        TODO: "",
        matter: "M-17",
        "[…]": "Final",
      }),
    );
    const out = path.join(dir, "filled.docx");
    const fill = captureIo();
    const code = await runFolioCli(
      [
        "placeholders",
        file,
        "--fill",
        `@${fields}`,
        "-o",
        out,
        "--expect-version",
        String(listed["fileVersion"]),
      ],
      fill.io,
    );
    expect(fill.stderr()).toBe("");
    expect(code).toBe(EXIT_CODES.ok);
    expect(dataOf(fill.stdout())["result"]).toMatchObject({
      filled: 7,
      unresolved: 0,
    });

    const check = captureIo();
    expect(await runFolioCli(["placeholders", out], check.io)).toBe(
      EXIT_CODES.ok,
    );
    expect(await readFile(file)).not.toEqual(await readFile(out));
  });

  test("a fill that changes nothing is refused", async () => {
    const file = path.join(dir, "draft.docx");
    await writeFile(file, await buildPackage(DRAFT));
    const io = captureIo();
    expect(
      await runFolioCli(
        [
          "placeholders",
          file,
          "--fill",
          '{"unknown":"value"}',
          "--in-place",
          "--no-expect-version",
        ],
        io.io,
      ),
    ).toBe(2);
    expect(envelopeOf(io.stdout())["error"]).toMatchObject({
      code: "operation_rejected",
    });
  });
});
