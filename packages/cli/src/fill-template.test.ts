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

const p = (id: string, text: string, properties = ""): string =>
  `<w:p w14:paraId="${id}" w14:textId="77777777">${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const style = (id: string): string => `<w:pStyle w:val="${id}"/>`;
const paragraphStyle = (id: string, extra = ""): string =>
  `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/><w:basedOn w:val="Normal"/>${extra}</w:style>`;

/** A filing template: front matter, an old body with a section break, and a signature block. */
const TEMPLATE: PackageFixture = {
  body:
    p("10000001", "IN THE MATTER OF {{party}}", style("Title")) +
    p("10000002", "Counsel: Old Firm") +
    p("10000003", "Old heading", style("Heading1")) +
    p(
      "10000004",
      "Old body text.",
      `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`,
    ) +
    p("10000005", "More old text.") +
    p("10000006", "Signed: ____________"),
  sectPr:
    '<w:sectPr><w:headerReference w:type="first" r:id="rIdFirst"/><w:footerReference w:type="default" r:id="rIdFooter"/>' +
    '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:titlePg/></w:sectPr>',
  headers: { rIdFirst: p("20000001", "Matter {{matter}}") },
  footers: { rIdFooter: p("30000001", "Footer stays") },
  styles:
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    paragraphStyle("Title") +
    paragraphStyle("Heading1", '<w:pPr><w:outlineLvl w:val="0"/></w:pPr>') +
    paragraphStyle("HeadingA", '<w:pPr><w:outlineLvl w:val="0"/></w:pPr>') +
    paragraphStyle("HeadingB", '<w:pPr><w:outlineLvl w:val="1"/></w:pPr>') +
    paragraphStyle("BodyNumbered") +
    paragraphStyle("Quote") +
    paragraphStyle("ListPara") +
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>',
  numbering:
    '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/></w:lvl></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="28"><w:abstractNumId w:val="1"/></w:num>' +
    '<w:num w:numId="5"><w:abstractNumId w:val="2"/></w:num>',
};

const STYLES = {
  h2: "HeadingA",
  h3: "HeadingB",
  numberedParagraph: { style: "BodyNumbered", numId: 5 },
  quote: "Quote",
  list: { style: "ListPara", numId: 28 },
  table: "TableGrid",
};

const MARKDOWN = [
  "## Facts",
  "",
  "The **claimant** relies on *two* points.",
  "",
  "1. First point.",
  "2. Second point.",
  "",
  "- A bullet",
  "  - A nested bullet",
  "",
  "> A quoted passage.",
  "",
  "| Exhibit | Page |",
  "| --- | --- |",
  "| A | 3 |",
  "",
  "### Relief",
  "",
  "The claimant asks for costs.",
].join("\n");

const fill = async (args: Record<string, unknown>) => {
  const outcome = (
    await runPackageRewrite("fillTemplate", await buildPackage(TEMPLATE), args)
  ).unwrap();
  const bytes = await outcome.bytes();
  if (bytes === null) throw new Error("expected a document");
  return {
    report: outcome.report,
    bytes,
    reviewer: await FolioDocxReviewer.fromBuffer(bytes.slice().buffer),
  };
};

const paragraphs = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.flatMap((block) =>
    block.kind === "diagnostic"
      ? []
      : [
          {
            text: block.text,
            ...(block.styleId !== undefined && { style: block.styleId }),
            ...(block.listReference !== undefined && {
              list: block.listReference,
            }),
          },
        ],
  );

describe("fill_template", () => {
  test("keeps the front matter and the signature, styles Markdown content through the map, and fills fields", async () => {
    const { report, bytes, reviewer } = await fill({
      markdown: MARKDOWN,
      styles: STYLES,
      keepFrom: "Signed:",
      fields: { party: "Acme Ltd", matter: "M-17" },
      replaceText: { "Old Firm": "New Chambers" },
    });
    expect(report).toEqual({
      kept: { front: 2, tail: 1 },
      removed: 3,
      inserted: 10,
      fields: { party: 1, matter: 1, "Old Firm": 1 },
      warnings: ["1 section break in the replaced body went with it."],
    });
    expect(paragraphs(reviewer)).toEqual([
      { text: "IN THE MATTER OF Acme Ltd", style: "Title" },
      { text: "Counsel: New Chambers" },
      { text: "Facts", style: "HeadingA" },
      { text: "The claimant relies on two points." },
      {
        text: "First point.",
        style: "BodyNumbered",
        list: { numId: 5, level: 0 },
      },
      {
        text: "Second point.",
        style: "BodyNumbered",
        list: { numId: 5, level: 0 },
      },
      { text: "A bullet", style: "ListPara", list: { numId: 28, level: 0 } },
      {
        text: "A nested bullet",
        style: "ListPara",
        list: { numId: 28, level: 1 },
      },
      { text: "A quoted passage.", style: "Quote" },
      { text: "Exhibit" },
      { text: "Page" },
      { text: "A" },
      { text: "3" },
      { text: "Relief", style: "HeadingB" },
      { text: "The claimant asks for costs." },
      { text: "Signed: ____________" },
    ]);
    const texts = await packageTexts(bytes);
    const document = texts.get("word/document.xml") ?? "";
    expect(document).toMatch(
      /<w:b\/><\/w:rPr><w:t xml:space="preserve">claimant/u,
    );
    expect(document).toMatch(/<w:i\/><\/w:rPr><w:t xml:space="preserve">two/u);
    expect(document).toContain('<w:tblStyle w:val="TableGrid"/>');
    expect(document).toContain("<w:tblHeader/>");
    expect(document).not.toContain("Old body text.");
    // The final section, its first-page header and footer stay.
    expect(document).toContain("<w:titlePg/>");
    expect(texts.get("word/header-rIdFirst.xml")).toContain("Matter M-17");
    expect(texts.get("word/footer-rIdFooter.xml")).toContain("Footer stays");
  });

  test("takes a block list and keeps the front matter through keepUntil", async () => {
    const { report, reviewer } = await fill({
      blocks: [
        { type: "heading", level: 2, text: "Grounds" },
        {
          type: "numberedParagraph",
          runs: ["The ", { text: "first", bold: true }, " ground."],
        },
        { type: "paragraph", text: "Plain.", style: "Quote" },
        { type: "table", rows: [["a", "b"]] },
      ],
      styles: STYLES,
      keepUntil: "10000003",
    });
    expect(report).toMatchObject({
      kept: { front: 3, tail: 0 },
      removed: 3,
      inserted: 4,
    });
    expect(paragraphs(reviewer).map(({ text }) => text)).toEqual([
      "IN THE MATTER OF {{party}}",
      "Counsel: Old Firm",
      "Old heading",
      "Grounds",
      "The first ground.",
      "Plain.",
      "a",
      "b",
    ]);
  });

  test("refuses content it cannot place", async () => {
    const template = await buildPackage(TEMPLATE);
    for (const [args, code] of [
      [{}, "invalid_input"],
      [{ markdown: "x", blocks: [] }, "invalid_input"],
      [{ markdown: "x", styles: { h2: "Missing" } }, "invalid_input"],
      [{ markdown: "x", styles: { list: { numId: 99 } } }, "invalid_input"],
      [{ markdown: "x", styles: { heading: "HeadingA" } }, "invalid_input"],
      [{ markdown: "x", keepUntil: "No such text" }, "not_found"],
      [{ blocks: [{ type: "image" }] }, "invalid_input"],
    ] as const) {
      const outcome = await runPackageRewrite("fillTemplate", template, args);
      expect(outcome.isErr() && outcome.error.code).toBe(code);
    }
    const headless = await buildPackage({
      ...TEMPLATE,
      body: p("10000001", "Only text"),
    });
    const outcome = await runPackageRewrite("fillTemplate", headless, {
      markdown: "x",
    });
    expect(outcome.isErr() && outcome.error.code).toBe("invalid_input");
  });
});

describe("folio fill", () => {
  let dir = "";
  let cleanup: () => Promise<void> = () => Promise.resolve();

  beforeEach(async () => {
    ({ dir, cleanup } = await makeTempDir());
  });

  afterEach(async () => {
    await cleanup();
  });

  test("reads content and maps from files and writes a new document", async () => {
    const template = path.join(dir, "template.docx");
    await writeFile(template, await buildPackage(TEMPLATE));
    const before = await readFile(template);
    await writeFile(path.join(dir, "content.md"), MARKDOWN);
    await writeFile(path.join(dir, "styles.json"), JSON.stringify(STYLES));
    await writeFile(
      path.join(dir, "fields.json"),
      JSON.stringify({ party: "Acme Ltd" }),
    );
    const out = path.join(dir, "filing.docx");
    const io = captureIo();
    const code = await runFolioCli(
      [
        "fill",
        template,
        "--markdown",
        `@${path.join(dir, "content.md")}`,
        "--styles",
        `@${path.join(dir, "styles.json")}`,
        "--fields",
        `@${path.join(dir, "fields.json")}`,
        "--keep-from",
        "Signed:",
        "-o",
        out,
        "--no-expect-version",
      ],
      io.io,
    );
    expect(io.stderr()).toBe("");
    expect(code).toBe(0);
    expect(dataOf(io.stdout())["result"]).toMatchObject({
      report: { inserted: 10 },
    });
    expect(await readFile(template)).toEqual(before);
    const reviewer = await FolioDocxReviewer.fromBuffer(
      new Uint8Array(await readFile(out)).buffer,
    );
    expect(reviewer.getContentAsText()).toContain(
      "The claimant asks for costs.",
    );
  });
});
