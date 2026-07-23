import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "../parser";
import { createEmptyDocx, repackDocx } from "../rezip";

const DOCUMENT_XML_PATH = "word/document.xml";

describe("paragraph indentation round-trip", () => {
  test("preserves an explicit zero first-line indent through a full save", async () => {
    const zip = await JSZip.loadAsync(await createEmptyDocx());
    zip.file(
      DOCUMENT_XML_PATH,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:ind w:firstLine="0"/></w:pPr>
      <w:r><w:t>Paragraph</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`,
    );
    const originalBuffer = await zip.generateAsync({ type: "arraybuffer" });
    const document = await parseDocx(originalBuffer, { preloadFonts: false });

    expect(document.package.document.content.at(0)?.formatting?.indentFirstLine).toBe(0);

    const saved = await repackDocx({ ...document, originalBuffer });
    const reparsed = await parseDocx(saved, { preloadFonts: false });

    expect(reparsed.package.document.content.at(0)?.formatting?.indentFirstLine).toBe(0);
  });
});
