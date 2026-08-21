import { validateDocumentModel, validateDocxPackage } from "@stll/docx-core";
import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { createStellaStyleDocumentPreset } from "../../style-sets/stellaStyle";
import type { HeaderFooter } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { parseDocx } from "../parser";
import { RELATIONSHIP_TYPES } from "../relsParser";
import { createDocx } from "../rezip";
import {
  bookmark,
  createTableOfContentsField,
  endnote,
  heading,
  type HeadingLevel,
  hyperlink,
  InvalidFolioReportBuilderOptionsError,
  pageBreak,
  paragraph,
  run,
  table,
} from "./build";

const LINK = "https://example.test/source";

const zipText = async (buffer: ArrayBuffer, path: string): Promise<string | undefined> =>
  JSZip.loadAsync(buffer).then((zip) => zip.file(path)?.async("string"));

const externalHyperlinkRel = (relsXml: string | undefined, href: string): boolean =>
  relsXml !== undefined &&
  [...relsXml.matchAll(/<Relationship\b[^>]*\/>/gu)].some(
    ({ 0: rel }) =>
      rel.includes(`Type="${RELATIONSHIP_TYPES.hyperlink}"`) &&
      rel.includes(`Target="${href}"`) &&
      rel.includes('TargetMode="External"'),
  );

const firstHyperlinkHref = (blocks: HeaderFooter["content"]): string | undefined => {
  for (const block of blocks) {
    if (block.type !== "paragraph") {
      continue;
    }
    const link = block.content.find((item) => item.type === "hyperlink");
    if (link?.type === "hyperlink") {
      return link.href;
    }
  }
  return undefined;
};

describe("hyperlink relationships outside the main body", () => {
  test("an endnote hyperlink is minted into endnotes.xml.rels and round-trips", async () => {
    const doc = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    const ref = endnote(doc, [paragraph([run("See "), hyperlink({ text: "source", href: LINK })])]);
    doc.package.document.content = [paragraph([run("Claim"), ref])];

    const buffer = await createDocx(doc);

    expect(externalHyperlinkRel(await zipText(buffer, "word/_rels/endnotes.xml.rels"), LINK)).toBe(
      true,
    );
    expect(externalHyperlinkRel(await zipText(buffer, "word/_rels/document.xml.rels"), LINK)).toBe(
      false,
    );

    const parsed = await parseDocx(buffer, { preloadFonts: false });
    expect(firstHyperlinkHref(parsed.package.endnotes?.at(0)?.content ?? [])).toBe(LINK);
  });

  test("a note part cased by the producer keeps its own rels part", async () => {
    const doc = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    const ref = endnote(doc, [paragraph([run("See "), hyperlink({ text: "source", href: LINK })])]);
    doc.package.document.content = [paragraph([run("Claim"), ref])];
    const zip = await JSZip.loadAsync(await createDocx(doc));

    // Re-case the note part and its rels the way some producers do.
    const rename = async (from: string, to: string) => {
      const xml = await zip.file(from)?.async("string");
      expect(xml).toBeDefined();
      zip.remove(from);
      zip.file(to, xml ?? "");
    };
    await rename("word/endnotes.xml", "word/Endnotes.xml");
    await rename("word/_rels/endnotes.xml.rels", "word/_rels/Endnotes.xml.rels");
    const recase = async (path: string) => {
      const xml = await zip.file(path)?.async("string");
      zip.file(path, (xml ?? "").replaceAll("endnotes.xml", "Endnotes.xml"));
    };
    await recase("[Content_Types].xml");
    await recase("word/_rels/document.xml.rels");

    const parsed = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
      preloadFonts: false,
    });
    const note = parsed.package.endnotes?.at(0);
    expect(note).toBeDefined();
    const secondLink = `${LINK}/second`;
    note?.content.push(paragraph([hyperlink({ text: "more", href: secondLink })]));

    const buffer = await createDocx(parsed);
    expect(await zipText(buffer, "word/_rels/endnotes.xml.rels")).toBeUndefined();
    expect(
      externalHyperlinkRel(await zipText(buffer, "word/_rels/Endnotes.xml.rels"), secondLink),
    ).toBe(true);
  });

  test("a footer hyperlink is minted into the footer's own rels", async () => {
    const doc = createEmptyDocument();
    const footerRId = "rIdFooter";
    doc.package.footers = new Map([
      [
        footerRId,
        {
          type: "footer",
          hdrFtrType: "default",
          content: [paragraph([hyperlink({ text: "Site", href: LINK })])],
        },
      ],
    ]);
    doc.package.document.finalSectionProperties.footerReferences = [
      { type: "default", rId: footerRId },
    ];

    const buffer = await createDocx(doc);
    expect(externalHyperlinkRel(await zipText(buffer, "word/_rels/footer1.xml.rels"), LINK)).toBe(
      true,
    );

    const parsed = await parseDocx(buffer, { preloadFonts: false });
    const footer = [...(parsed.package.footers?.values() ?? [])].at(0);
    expect(firstHyperlinkHref(footer?.content ?? [])).toBe(LINK);
  });
});

describe("from-scratch header materialisation", () => {
  test("a header referenced from section properties becomes a real part", async () => {
    const doc = createEmptyDocument();
    // The seed rels already use rId1 for styles; the header must not collide.
    const headerRId = "rId1";
    doc.package.headers = new Map([
      [headerRId, { type: "header", hdrFtrType: "default", content: [paragraph("Report")] }],
    ]);
    doc.package.document.finalSectionProperties.headerReferences = [
      { type: "default", rId: headerRId },
    ];

    const buffer = await createDocx(doc);
    const zip = await JSZip.loadAsync(buffer);
    expect(zip.file("word/header1.xml")).not.toBeNull();

    const relsXml = await zip.file("word/_rels/document.xml.rels")!.async("string");
    const ids = [...relsXml.matchAll(/Id="(?<id>[^"]+)"/gu)].map((match) => match.groups!.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(relsXml).toContain(`Type="${RELATIONSHIP_TYPES.header}" Target="header1.xml"`);

    const parsed = await parseDocx(buffer, { preloadFonts: false });
    const header = [...(parsed.package.headers?.values() ?? [])].at(0);
    expect(header?.content.at(0)?.type).toBe("paragraph");
    expect(parsed.package.document.finalSectionProperties.headerReferences).toHaveLength(1);
  });
});

describe("table of contents field", () => {
  test("serialises a dirty TOC field and the supporting styles", async () => {
    const doc = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    doc.package.settings = { ...doc.package.settings, defaultTabStop: 720, updateFields: true };
    doc.package.document.content = [
      createTableOfContentsField(),
      heading({ text: "One", level: 1 }),
      heading({ text: "Two", level: 2 }),
      heading({ text: "Three", level: 3 }),
    ];

    const buffer = await createDocx(doc);
    const documentXml = await zipText(buffer, "word/document.xml");
    expect(documentXml).toContain('<w:fldChar w:fldCharType="begin" w:dirty="true"/>');
    expect(await zipText(buffer, "word/settings.xml")).toContain('<w:updateFields w:val="true"/>');

    const stylesXml = (await zipText(buffer, "word/styles.xml")) ?? "";
    for (const styleId of [
      "TOCHeading",
      "TOC1",
      "TOC2",
      "TOC3",
      "Heading1",
      "Heading6",
      "EndnoteReference",
      "EndnoteText",
    ]) {
      expect(stylesXml).toContain(`w:styleId="${styleId}"`);
    }

    const parsed = await parseDocx(buffer, { preloadFonts: false });
    const first = parsed.package.document.content.at(0);
    const field = first?.type === "paragraph" ? first.content.at(0) : undefined;
    expect(field?.type).toBe("complexField");
    if (field?.type !== "complexField") {
      return;
    }
    expect(field.fieldType).toBe("TOC");
    expect(field.instruction).toBe('TOC \\o "1-3" \\h \\z \\u');
    expect(field.dirty).toBe(true);
    expect(parsed.package.settings?.updateFields).toBe(true);
    const heading1 = parsed.package.styles?.styles.find((style) => style.styleId === "Heading1");
    expect(heading1?.pPr?.outlineLevel).toBe(0);
  });

  test("honours level and hyperlink options", () => {
    const field = createTableOfContentsField({
      levels: { from: 1, to: 2 },
      hyperlinks: false,
    }).content.at(0);
    expect(field?.type === "complexField" && field.instruction).toBe('TOC \\o "1-2" \\z \\u');
  });
});

describe("builder input validation", () => {
  const invalidTocLevels = [
    { from: 0, to: 3 },
    { from: 1, to: 10 },
    { from: 3, to: 1 },
    { from: 1.5, to: 3 },
    { from: Number.NaN, to: 3 },
  ];
  test.each(invalidTocLevels)("rejects TOC levels %o", (levels) => {
    expect(() => createTableOfContentsField({ levels })).toThrow(
      InvalidFolioReportBuilderOptionsError,
    );
  });

  test("accepts the full 1-9 TOC range", () => {
    expect(() => createTableOfContentsField({ levels: { from: 1, to: 9 } })).not.toThrow();
  });

  test.each([0, -1, 1.5, Number.NaN])("rejects gridSpan %p", (gridSpan) => {
    expect(() => table({ rows: [[{ content: [paragraph("x")], gridSpan }]] })).toThrow(
      InvalidFolioReportBuilderOptionsError,
    );
  });

  test.each([0, -100, 0.5, Number.POSITIVE_INFINITY])("rejects column width %p", (width) => {
    expect(() => table({ rows: [["a", "b"]], columnWidths: [1000, width] })).toThrow(
      InvalidFolioReportBuilderOptionsError,
    );
  });

  test("column widths are validated even when no row uses them", () => {
    expect(() => table({ rows: [], columnWidths: [0] })).toThrow(
      InvalidFolioReportBuilderOptionsError,
    );
  });

  test("rejects a heading level outside HEADING_LEVELS", () => {
    // SAFETY: test deliberately bypasses the static type to exercise the runtime guard
    const level = 7 as HeadingLevel;
    expect(() => heading({ text: "x", level })).toThrow(InvalidFolioReportBuilderOptionsError);
  });

  test("the error carries the offending path", () => {
    expect(() => table({ rows: [["a"]], columnWidths: [1000, -1] })).toThrow(
      expect.objectContaining({ path: "columnWidths[1]" }),
    );
  });
});

describe("report builders end to end", () => {
  test("builds a report that round-trips and validates", async () => {
    const doc = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    const shade = { fill: { rgb: "D9E2F3" }, pattern: "clear" as const };
    const ref = endnote(doc, [
      paragraph([run("Methodology: "), hyperlink({ text: "public record", href: LINK })]),
    ]);
    doc.package.document.content = [
      paragraph(bookmark({ name: "cover", content: [run("Annual findings", { bold: true })] }), {
        styleId: "Title",
      }),
      createTableOfContentsField(),
      pageBreak(),
      heading({ text: "Summary", level: 1 }),
      paragraph([run("See the "), hyperlink({ text: "cover", anchor: "cover" }), run("."), ref]),
      heading({ text: "Findings", level: 2 }),
      table({
        header: ["Area", "Finding", "Severity"],
        headerShading: shade,
        columnWidths: [2400, 4800, 1800],
        rows: [
          [{ content: [paragraph("Contracts")], vMerge: "restart" }, "Missing signature", "High"],
          [{ content: [], vMerge: "continue" }, "Expired term", "Low"],
          [{ content: [paragraph("Total")], gridSpan: 2 }, "2"],
        ],
      }),
    ];

    const model = validateDocumentModel(doc);
    expect(model.issues.filter((issue) => issue.severity === "error")).toEqual([]);

    const buffer = await createDocx(doc);
    expect(await validateDocxPackage(buffer)).toEqual({ valid: true });

    const parsed = await parseDocx(buffer, { preloadFonts: false });
    expect(validateDocumentModel(parsed).issues.filter((i) => i.severity === "error")).toEqual([]);

    const blocks = parsed.package.document.content;
    const title = blocks.at(0);
    expect(
      title?.type === "paragraph" && title.content.some((c) => c.type === "bookmarkStart"),
    ).toBe(true);
    const headings = blocks.flatMap((block) =>
      block.type === "paragraph" && block.formatting?.styleId?.startsWith("Heading")
        ? [block.formatting.styleId]
        : [],
    );
    expect(headings).toEqual(["Heading1", "Heading2"]);

    const parsedTable = blocks.find((block) => block.type === "table");
    expect(parsedTable?.type).toBe("table");
    if (parsedTable?.type !== "table") {
      return;
    }
    expect(parsedTable.rows).toHaveLength(4);
    expect(parsedTable.rows[0]?.formatting?.header).toBe(true);
    expect(parsedTable.rows[0]?.cells[0]?.formatting?.shading?.fill?.rgb).toBe("D9E2F3");
    expect(parsedTable.rows[1]?.cells[0]?.formatting?.vMerge).toBe("restart");
    expect(parsedTable.rows[2]?.cells[0]?.formatting?.vMerge).toBe("continue");
    expect(parsedTable.rows[3]?.cells[0]?.formatting?.gridSpan).toBe(2);

    expect(parsed.package.endnotes).toHaveLength(1);
    expect(firstHyperlinkHref(parsed.package.endnotes?.at(0)?.content ?? [])).toBe(LINK);
  });
});
