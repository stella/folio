import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { docxToMarkdown } from "./docxToMarkdown";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const p = (text: string, style?: string) =>
  `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}<w:r><w:t>${text}</w:t></w:r></w:p>`;

const makeDocx = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "word/document.xml",
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}</w:body></w:document>`,
  );
  return await zip.generateAsync({ type: "uint8array" });
};

describe("docxToMarkdown", () => {
  test("composes parseDocx + toMarkdown and keeps content-control text at every position", async () => {
    // Heading + a block-level content control + a table whose second cell wraps
    // its value in a content control. The hand-rolled OOXML walker this
    // replaces dropped the table-cell control; folio's parser must not.
    const bytes = await makeDocx(
      p("Agreement", "Heading1") +
        `<w:sdt><w:sdtContent>${p("Controlled clause text.")}</w:sdtContent></w:sdt>` +
        `<w:tbl><w:tr>` +
        `<w:tc>${p("Party")}</w:tc>` +
        `<w:tc><w:sdt><w:sdtContent>${p("Acme Corp")}</w:sdtContent></w:sdt></w:tc>` +
        `</w:tr></w:tbl>`,
    );

    const markdown = await docxToMarkdown(bytes);

    expect(markdown).toContain("Agreement");
    expect(markdown).toContain("Controlled clause text.");
    expect(markdown).toContain("Party");
    expect(markdown).toContain("Acme Corp");
  });
});
