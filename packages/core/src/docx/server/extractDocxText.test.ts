import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { RELATIONSHIP_TYPES } from "../relsParser";
import { extractDocxText } from "./extractDocxText";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

type MakeDocxOptions = {
  body: string;
  headers?: Record<string, string>;
  footers?: Record<string, string>;
  /**
   * Header/footer part paths present in the archive but deliberately left
   * out of `word/_rels/document.xml.rels` + every section's
   * `w:headerReference` / `w:footerReference` — an orphan part no section
   * actually wires up, the way a stale or planted part would look.
   */
  orphanParts?: Record<string, "header" | "footer">;
};

/**
 * Build a minimal DOCX. When `headers`/`footers` are given, also wires them
 * up the way a real Word document does — a relationship in
 * `word/_rels/document.xml.rels` plus a matching `w:headerReference` /
 * `w:footerReference` on the body's `w:sectPr` — since `extractDocxText`
 * resolves referenced parts through that wiring rather than by filename.
 */
const makeDocx = async ({
  body,
  headers = {},
  footers = {},
  orphanParts = {},
}: MakeDocxOptions): Promise<Uint8Array> => {
  const zip = new JSZip();
  const relationships: string[] = [];
  const sectionReferences: string[] = [];
  let nextRId = 1;

  const wireParts = (
    parts: Record<string, string>,
    kind: "header" | "footer",
    relationshipType: string,
  ): void => {
    const rootName = kind === "header" ? "hdr" : "ftr";
    for (const [path, content] of Object.entries(parts)) {
      zip.file(path, `<w:${rootName} xmlns:w="${W_NS}">${content}</w:${rootName}>`);
      const rId = `rId${nextRId++}`;
      const target = path.replace(/^word\//, "");
      relationships.push(
        `<Relationship Id="${rId}" Type="${relationshipType}" Target="${target}"/>`,
      );
      sectionReferences.push(`<w:${kind}Reference w:type="default" r:id="${rId}"/>`);
    }
  };

  wireParts(headers, "header", RELATIONSHIP_TYPES.header);
  wireParts(footers, "footer", RELATIONSHIP_TYPES.footer);

  for (const [path, kind] of Object.entries(orphanParts)) {
    const rootName = kind === "header" ? "hdr" : "ftr";
    zip.file(path, `<w:${rootName} xmlns:w="${W_NS}">${paragraph("Orphan")}</w:${rootName}>`);
    // Deliberately no Relationship entry and no section reference — this
    // part exists in the archive but nothing wires it up.
  }

  const sectPr =
    sectionReferences.length > 0 ? `<w:sectPr>${sectionReferences.join("")}</w:sectPr>` : "";
  zip.file(
    "word/document.xml",
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}${sectPr}</w:body></w:document>`,
  );
  if (relationships.length > 0) {
    zip.file(
      "word/_rels/document.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NS}">${relationships.join("")}</Relationships>`,
    );
  }
  return await zip.generateAsync({ type: "uint8array" });
};

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

const table = (rows: string, properties = "") => `<w:tbl>${properties}${rows}</w:tbl>`;
const row = (cells: string, properties = "") => `<w:tr>${properties}${cells}</w:tr>`;
const cell = (content: string, properties = "") =>
  `<w:tc><w:tcPr>${properties}</w:tcPr>${content}</w:tc>`;

/** `w:tblPr` as Word writes it on every table, header row or not. */
const DEFAULT_TABLE_PROPERTIES = `<w:tblPr><w:tblLook w:firstColumn="1" w:firstRow="1" w:lastRow="0" w:val="04A0"/></w:tblPr>`;

/** `w:trPr` for a row Word repeats at the top of each page: OOXML's header row. */
const HEADER_ROW_PROPERTIES = `<w:trPr><w:tblHeader/></w:trPr>`;

/**
 * Cells in one rendered row. Escaped pipes are HTML entities, so every
 * remaining `|` between the outer delimiters separates two columns.
 */
const columnCount = (line: string): number => line.slice(1, -1).split("|").length;

describe("extractDocxText", () => {
  test("returns deterministic paragraphs across document parts", async () => {
    const bytes = await makeDocx({
      body:
        paragraph("Before") +
        `<w:tbl><w:tr><w:tc>${paragraph("Cell")}</w:tc></w:tr></w:tbl>` +
        `<w:sdt><w:sdtContent>${paragraph("Control")}</w:sdtContent></w:sdt>` +
        paragraph("After"),
      headers: {
        "word/header2.xml": paragraph("Header 2"),
        "word/header1.xml": paragraph("Header 1"),
      },
      footers: {
        "word/footer1.xml": paragraph("Footer"),
      },
    });

    const result = await extractDocxText(bytes);

    expect(
      result.paragraphs.map(({ index, text, source }) => ({
        index,
        text,
        source,
      })),
    ).toEqual([
      { index: 0, text: "Header 1", source: "header" },
      { index: 1, text: "Header 2", source: "header" },
      { index: 2, text: "Before", source: "body" },
      { index: 3, text: "|  |", source: "body" },
      { index: 4, text: "| --- |", source: "body" },
      { index: 5, text: "| Cell |", source: "body" },
      { index: 6, text: "Control", source: "body" },
      { index: 7, text: "After", source: "body" },
      { index: 8, text: "Footer", source: "footer" },
    ]);
    expect(result.charCount).toBe(
      result.paragraphs.reduce((sum, { text }) => sum + text.length, 0),
    );
    expect(result.view).toBe("accepted");
  });

  test("excludes an orphaned header/footer part no section references", async () => {
    const bytes = await makeDocx({
      body: paragraph("Before"),
      headers: {
        "word/header1.xml": paragraph("Referenced header"),
      },
      orphanParts: {
        // Present in the archive, and filename-shaped like a real header,
        // but wired into neither the rels nor any section — stale content
        // (or a planted prompt-injection payload) that Word itself never
        // renders must not be surfaced here either.
        "word/header2.xml": "header",
        "word/footer1.xml": "footer",
      },
    });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.map(({ text, source }) => ({ text, source }))).toEqual([
      { text: "Referenced header", source: "header" },
      { text: "Before", source: "body" },
    ]);
  });

  test("preserves explicit whitespace and applies the accepted revision view", async () => {
    const bytes = await makeDocx({
      body: `<w:p>
        <w:r><w:t>Keep</w:t><w:br/><w:t>Line</w:t><w:tab/><w:t>Tab</w:t></w:r>
        <w:del><w:r><w:delText>Deleted</w:delText></w:r></w:del>
        <w:moveFrom><w:r><w:t>Moved away</w:t></w:r></w:moveFrom>
        <w:ins><w:r><w:t>Inserted</w:t></w:r></w:ins>
      </w:p>`,
    });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.at(0)?.text).toBe("Keep\nLine\tTabInserted");
  });

  test("extracts paragraph and majority-run formatting metadata", async () => {
    const bytes = await makeDocx({
      body: `<w:p>
        <w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>
        <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Majority</w:t></w:r>
        <w:r><w:t>x</w:t></w:r>
      </w:p>
      <w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>Left</w:t></w:r></w:p>`,
    });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.at(0)).toEqual({
      index: 0,
      text: "Majorityx",
      source: "body",
      style: "Heading1",
      alignment: "center",
      bold: true,
      fontSize: 28,
    });
    expect(result.paragraphs.at(1)?.alignment).toBe("left");
  });

  test("supports alternate prefixes for the main OOXML namespace", async () => {
    const zip = new JSZip();
    zip.file(
      "word/document.xml",
      `<x:document xmlns:x="${W_NS}"><x:body><x:p><x:pPr><x:pStyle x:val="Title"/></x:pPr><x:r><x:t>Alternate</x:t></x:r></x:p></x:body></x:document>`,
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.at(0)?.text).toBe("Alternate");
    expect(result.paragraphs.at(0)?.style).toBe("Title");
  });

  test("emits table cells as markdown rows, not a flat row-major paragraph list", async () => {
    // Shaped like an SEC cover page: a row of values over a row of labels.
    // Flattened row-major, nothing ties an EIN to the column that names it.
    const bytes = await makeDocx({
      body: table(
        row(cell(paragraph("Delaware")) + cell(paragraph("4911")) + cell(paragraph("83-4027615"))) +
          row(
            cell(paragraph("(State or Other Jurisdiction of Incorporation or Organization)")) +
              cell(paragraph("(Primary Standard Industrial Classification Code Number)")) +
              cell(paragraph("(I.R.S. Employer Identification Number)")),
          ),
        DEFAULT_TABLE_PROPERTIES,
      ),
    });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.map(({ text }) => text)).toEqual([
      "|  |  |  |",
      "| --- | --- | --- |",
      "| Delaware | 4911 | 83-4027615 |",
      "| (State or Other Jurisdiction of Incorporation or Organization) | (Primary Standard Industrial Classification Code Number) | (I.R.S. Employer Identification Number) |",
    ]);
  });

  test("marks table rows with their table and role, and leaves prose unmarked", async () => {
    const oneByOne = table(row(cell(paragraph("x"))));
    const bytes = await makeDocx({
      body: paragraph("Prose") + oneByOne + oneByOne,
      headers: { "word/header1.xml": oneByOne },
    });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.map(({ text, tableRow }) => [text, tableRow])).toEqual([
      ["|  |", { table: 0, kind: "syntheticHeader" }],
      ["| --- |", { table: 0, kind: "delimiter" }],
      ["| x |", { table: 0, kind: "cells" }],
      ["Prose", undefined],
      ["|  |", { table: 1, kind: "syntheticHeader" }],
      ["| --- |", { table: 1, kind: "delimiter" }],
      ["| x |", { table: 1, kind: "cells" }],
      ["|  |", { table: 2, kind: "syntheticHeader" }],
      ["| --- |", { table: 2, kind: "delimiter" }],
      ["| x |", { table: 2, kind: "cells" }],
    ]);
  });

  test("promotes the first row only where the document marks it as a header row", async () => {
    const cells = cell(paragraph("H1")) + cell(paragraph("H2"));
    const data = row(cell(paragraph("v1")) + cell(paragraph("v2")));

    const declared = await extractDocxText(
      await makeDocx({
        body: table(row(cells, HEADER_ROW_PROPERTIES) + data, DEFAULT_TABLE_PROPERTIES),
      }),
    );
    // `w:tblLook/@w:firstRow` is conditional formatting Word writes on every
    // table, so on its own it must not turn a row of data into column names.
    const undeclared = await extractDocxText(
      await makeDocx({ body: table(row(cells) + data, DEFAULT_TABLE_PROPERTIES) }),
    );

    expect(declared.paragraphs.map(({ text }) => text)).toEqual([
      "| H1 | H2 |",
      "| --- | --- |",
      "| v1 | v2 |",
    ]);
    expect(undeclared.paragraphs.map(({ text }) => text)).toEqual([
      "|  |  |",
      "| --- | --- |",
      "| H1 | H2 |",
      "| v1 | v2 |",
    ]);
    expect(declared.paragraphs.at(0)?.tableRow).toEqual({ table: 0, kind: "cells" });
    expect(undeclared.paragraphs.at(0)?.tableRow).toEqual({ table: 0, kind: "syntheticHeader" });
  });

  test("flattens a horizontally merged cell across the grid columns it spans", async () => {
    const bytes = await makeDocx({
      body: table(
        row(cell(paragraph("Spans two"), `<w:gridSpan w:val="2"/>`) + cell(paragraph("C"))) +
          row(cell(paragraph("a1")) + cell(paragraph("a2")) + cell(paragraph("a3"))),
      ),
    });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.map(({ text }) => text)).toEqual([
      "|  |  |  |",
      "| --- | --- | --- |",
      "| Spans two |  | C |",
      "| a1 | a2 | a3 |",
    ]);
  });

  test("renders a vertical-merge continuation as an empty column", async () => {
    const bytes = await makeDocx({
      body: table(
        row(cell(paragraph("Merged down"), `<w:vMerge w:val="restart"/>`) + cell(paragraph("r1"))) +
          // Word renders no content for a continuation cell. Content left here
          // by a producer must not surface, and the anchor must not be repeated:
          // either would invent a fact the document does not show.
          row(cell(paragraph("Stale"), `<w:vMerge/>`) + cell(paragraph("r2"))),
      ),
    });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.map(({ text }) => text)).toEqual([
      "|  |  |",
      "| --- | --- |",
      "| Merged down | r1 |",
      "|  | r2 |",
    ]);
  });

  test("flattens a nested table into the cell that contains it", async () => {
    const bytes = await makeDocx({
      body: table(
        row(
          cell(paragraph("outer")) +
            cell(
              table(
                row(cell(paragraph("in1")) + cell(paragraph("in2"))) +
                  row(cell(paragraph("in3")) + cell(paragraph("in4"))),
              ),
            ),
        ),
      ),
    });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.map(({ text }) => text)).toEqual([
      "|  |  |",
      "| --- | --- |",
      "| outer | in1 / in2<br>in3 / in4 |",
    ]);
  });

  test("keeps a multi-paragraph cell on one row and escapes what would split it", async () => {
    const bytes = await makeDocx({
      body: table(
        row(
          cell(paragraph("first") + `<w:p/>` + paragraph("second")) +
            cell("") +
            cell(paragraph("pipe|inside")) +
            cell(`<w:p><w:r><w:t>break</w:t><w:br/><w:t>after</w:t></w:r></w:p>`),
        ),
      ),
    });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.map(({ text }) => text)).toEqual([
      "|  |  |  |  |",
      "| --- | --- | --- | --- |",
      "| first<br>second |  | pipe&#124;inside | break<br>after |",
    ]);
  });

  test("gives every row of a table the same column count", async () => {
    const bytes = await makeDocx({
      body: table(
        row(cell(paragraph("wide"), `<w:gridSpan w:val="3"/>`)) +
          row(cell(paragraph("a")) + cell(paragraph("b"))) +
          row(cell(paragraph("only"))) +
          row(""),
      ),
    });

    const result = await extractDocxText(bytes);
    const counts = result.paragraphs.map(({ text }) => columnCount(text));

    expect(new Set(counts)).toEqual(new Set([3]));
    expect(result.paragraphs.map(({ text }) => text)).toEqual([
      "|  |  |  |",
      "| --- | --- | --- |",
      "| wide |  |  |",
      "| a | b |  |",
      "| only |  |  |",
      "|  |  |  |",
    ]);
  });

  test("emits nothing for a table that has no cell at all", async () => {
    const bytes = await makeDocx({
      body: paragraph("Before") + table("") + table(row("")) + paragraph("After"),
    });

    const result = await extractDocxText(bytes);

    expect(result.paragraphs.map(({ text }) => text)).toEqual(["Before", "After"]);
  });

  test("counts every emitted paragraph's text, table rows included", async () => {
    const bytes = await makeDocx({
      body:
        paragraph("Prose") +
        table(
          row(cell(paragraph("a")) + cell(paragraph("b")), HEADER_ROW_PROPERTIES),
          DEFAULT_TABLE_PROPERTIES,
        ),
      headers: { "word/header1.xml": table(row(cell(paragraph("h")))) },
      footers: { "word/footer1.xml": paragraph("Footer") },
    });

    const result = await extractDocxText(bytes);

    expect(result.charCount).toBe(
      result.paragraphs.reduce((sum, { text }) => sum + text.length, 0),
    );
  });

  test("returns an empty result when the main document part is absent", async () => {
    const zip = new JSZip();
    zip.file("other.xml", "<root/>");
    const bytes = await zip.generateAsync({ type: "uint8array" });

    expect(await extractDocxText(bytes)).toEqual({
      paragraphs: [],
      charCount: 0,
      view: "accepted",
    });
  });
});
