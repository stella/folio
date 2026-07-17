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
      relationships.push(`<Relationship Id="${rId}" Type="${relationshipType}" Target="${target}"/>`);
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

  const sectPr = sectionReferences.length > 0 ? `<w:sectPr>${sectionReferences.join("")}</w:sectPr>` : "";
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
      { index: 3, text: "Cell", source: "body" },
      { index: 4, text: "Control", source: "body" },
      { index: 5, text: "After", source: "body" },
      { index: 6, text: "Footer", source: "footer" },
    ]);
    expect(result.charCount).toBe("Header 1Header 2BeforeCellControlAfterFooter".length);
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
