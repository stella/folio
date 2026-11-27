import { beforeAll, describe, expect, test } from "bun:test";

import {
  buildTemplate,
  fieldRuns,
  NS,
  p,
  r,
  readPart,
  replacePart,
  XML,
} from "./fixtures";
import {
  applyFor,
  buildSourceWith,
  issuesOf,
  pictureRun,
  PIXEL_PNG,
  planFor,
  unwrap,
  verifyFor,
} from "./planning";

const header = (inner: string): string => `${XML}<w:hdr ${NS}>${inner}</w:hdr>`;
const footer = (inner: string): string => `${XML}<w:ftr ${NS}>${inner}</w:ftr>`;

const DISCLAIMER =
  "This sample document is a starting point only and must be tailored.";

/** A source whose first section shows a first-page header (a notice, a
 *  logo, a table), a header for the other pages (a draft line and a page
 *  number) and a footer (a version line), and a comment mark in a header. */
const sourceWithHeaders = async (
  { titlePage }: { titlePage: boolean } = { titlePage: true },
): Promise<Uint8Array> =>
  await buildSourceWith({
    body: [
      p(
        r("The parties agree as follows."),
        `<w:sectPr><w:headerReference w:type="first" r:id="rId20"/><w:headerReference w:type="default" r:id="rId21"/><w:footerReference w:type="default" r:id="rId22"/>${titlePage ? "<w:titlePg/>" : ""}</w:sectPr>`,
      ),
      p(r("Each party signs below.")),
    ].join(""),
    relationships: [
      { id: "rId20", type: "header", target: "header2.xml" },
      { id: "rId21", type: "header", target: "header3.xml" },
      { id: "rId22", type: "footer", target: "footer2.xml" },
    ],
    parts: {
      "word/header2.xml": header(
        `${p(r(DISCLAIMER, "<w:b/>"))}${p(pictureRun("rId5"))}<w:tbl><w:tr><w:tc>${p(r("Letterhead cell"))}</w:tc></w:tr></w:tbl>`,
      ),
      "word/header3.xml": header(
        `${p(`<w:bookmarkStart w:id="9" w:name="top"/>${r("Draft for discussion")}<w:bookmarkEnd w:id="9"/>`, '<w:jc w:val="right"/>')}${p(`${r("Page ")}${fieldRuns("PAGE", r("2"))}`)}`,
      ),
      "word/footer2.xml": footer(p(r("Version 3, last updated in November"))),
      "word/_rels/header2.xml.rels": `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/></Relationships>`,
      "word/media/logo.png": PIXEL_PNG,
    },
  });

let template: Uint8Array;
let titledTemplate: Uint8Array;

beforeAll(async () => {
  template = await buildTemplate();
  const document = (await readPart(template, "word/document.xml")) ?? "";
  titledTemplate = await replacePart(
    template,
    "word/document.xml",
    document.replace("<w:pgSz ", "<w:titlePg/><w:pgSz "),
  );
});

describe("the source's headers and footers", () => {
  test("are left to the template's unless the plan keeps them", async () => {
    const applied = unwrap(
      await applyFor({ source: await sourceWithHeaders(), template }),
    );
    expect(applied.report).not.toHaveProperty("sourceHeaders");
    expect(await readPart(applied.docx, "word/header1.xml")).not.toContain(
      "Draft for discussion",
    );
  });

  test("go after the template's for the same pages, page numbers, pictures and tables included", async () => {
    const applied = unwrap(
      await applyFor({
        source: await sourceWithHeaders(),
        template: titledTemplate,
        extra: { sourceHeaders: "keep" },
      }),
    );
    const report = applied.report.sourceHeaders;
    expect(
      report?.kept.map(({ family, kind, created }) => [family, kind, created]),
    ).toEqual([
      ["header", "first", true],
      ["header", "default", false],
      ["footer", "default", false],
    ]);
    expect(report?.leftOut).toEqual({
      firstPageParts: 0,
      evenPageParts: 0,
      sharedParts: 0,
      unsupportedObjects: 0,
    });
    const other = (await readPart(applied.docx, "word/header1.xml")) ?? "";
    expect(other.indexOf("House letterhead")).toBeLessThan(
      other.indexOf("Draft for discussion"),
    );
    // The page number stays a field; the bookmark is the source's own.
    expect(other.match(/PAGE/gu)?.length).toBe(2);
    expect(other).not.toContain("bookmarkStart");
    const first = report?.kept.find(({ kind }) => kind === "first")?.part;
    const firstXml = (await readPart(applied.docx, first ?? "")) ?? "";
    expect(firstXml).toContain(DISCLAIMER);
    expect(firstXml).toContain("Letterhead cell");
    expect(firstXml).toContain("<w:drawing>");
    const relationships =
      (await readPart(
        applied.docx,
        (first ?? "").replace("word/", "word/_rels/") + ".rels",
      )) ?? "";
    expect(relationships).toContain("media/placed");
    expect((await readPart(applied.docx, "word/footer1.xml")) ?? "").toContain(
      "Version 3, last updated in November",
    );
    const types = (await readPart(applied.docx, "[Content_Types].xml")) ?? "";
    expect(types).toContain(first ?? "?");
  });

  test("a first page the output does not set apart leaves the source's first-page header out, and says so", async () => {
    const applied = unwrap(
      await applyFor({
        source: await sourceWithHeaders(),
        template,
        extra: { sourceHeaders: "keep" },
      }),
    );
    expect(applied.report.sourceHeaders?.leftOut.firstPageParts).toBe(1);
    expect(
      applied.report.sourceHeaders?.kept.map(({ kind }) => kind),
    ).not.toContain("first");
  });

  test("a source without a distinct first page shows its other header there too", async () => {
    const applied = unwrap(
      await applyFor({
        source: await sourceWithHeaders({ titlePage: false }),
        template: titledTemplate,
        extra: { sourceHeaders: "keep" },
      }),
    );
    const kinds = applied.report.sourceHeaders?.kept.flatMap(
      ({ family, kind }) => (family === "header" ? [kind] : []),
    );
    expect(kinds).toEqual(["first", "default"]);
  });

  test("a carried block brings the namespaces of its part and leaves the source's styles and lists behind", async () => {
    const WP =
      "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    const source = await buildSourceWith({
      body: p(
        r("Text."),
        '<w:sectPr><w:headerReference w:type="default" r:id="rId20"/></w:sectPr>',
      ),
      relationships: [{ id: "rId20", type: "header", target: "header2.xml" }],
      parts: {
        "word/header2.xml": `${XML}<w:hdr ${NS} xmlns:wp="${WP}">${p(
          `${r("Notice")}<w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="4" name="Shape"/></wp:inline></w:drawing></w:r>`,
          '<w:pStyle w:val="SourceHeader"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr>',
        )}</w:hdr>`,
      },
    });
    const applied = unwrap(
      await applyFor({
        source,
        template: titledTemplate,
        extra: { sourceHeaders: "keep" },
      }),
    );
    const part = applied.report.sourceHeaders?.kept.at(0)?.part ?? "";
    const xml = (await readPart(applied.docx, part)) ?? "";
    expect(xml).toContain(`xmlns:wp="${WP}"`);
    expect(xml).not.toContain("SourceHeader");
    expect(xml).not.toContain("numPr");
    expect(xml).toContain("Notice");
  });

  test("an output whose header lost the carried content fails verification", async () => {
    const source = await sourceWithHeaders();
    const plan = await planFor({
      source,
      template,
      extra: { sourceHeaders: "keep" },
    });
    const applied = unwrap(
      await applyFor({ source, template, extra: { sourceHeaders: "keep" } }),
    );
    const xml = (await readPart(applied.docx, "word/header1.xml")) ?? "";
    const tampered = await replacePart(
      applied.docx,
      "word/header1.xml",
      xml.replace("Draft for discussion", "Final"),
    );
    const report = await verifyFor({
      source,
      template,
      output: tampered,
      plan,
    });
    expect(report.ok).toBe(false);
    expect(report.issues.map(({ code }) => code)).toContain("header-mismatch");
    expect(issuesOf(await applyFor({ source, template }))).toEqual([]);
  });
});
