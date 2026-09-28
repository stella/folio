import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`;
const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`;
const NUMBERING_XML = `<w:numbering xmlns:w="${W}" xmlns:ve="${MC}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

const documentXml = (nestedRebinding: boolean): string =>
  `<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:ve="${MC}"><w:body>` +
  `<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>Beta</w:t></w:r></w:p>` +
  `<w:p><w:r${nestedRebinding ? ' xmlns:w14="urn:other" xmlns:ve="urn:other"' : ""}><w:t>Gamma</w:t></w:r></w:p>` +
  `<w:sectPr/></w:body></w:document>`;

const packageWith = async (xml: string): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", ROOT_RELS);
  zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS);
  zip.file("word/document.xml", xml);
  zip.file("word/numbering.xml", NUMBERING_XML);
  const bytes = await zip.generateAsync({ type: "uint8array" });
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};

describe("selective save with an aliased markup-compatibility prefix", () => {
  for (const nestedRebinding of [false, true]) {
    test(`one-character edit stays local with nested rebinding ${nestedRebinding}`, async () => {
      const sourceXml = documentXml(nestedRebinding);
      const reviewer = await FolioDocxReviewer.fromBuffer(await packageWith(sourceXml));
      const target = reviewer.snapshot().blocks.find((block) => block.text === "Beta");
      if (!target) throw new Error("Target paragraph missing");
      const result = reviewer.applyOperations(
        [{ id: "edit", type: "replaceInBlock", blockId: target.id, find: "Beta", replace: "Beto" }],
        { mode: "direct" },
      );
      expect(result.applied).toHaveLength(1);

      const saved = await JSZip.loadAsync(await reviewer.toBuffer());
      const savedXml = await saved.file("word/document.xml")?.async("text");
      expect(savedXml).toBeDefined();
      expect(savedXml).toContain("Beto");
      expect(savedXml).not.toContain("w14:paraId=");
      const targetStart = sourceXml.indexOf("<w:p><w:r><w:t>Beta");
      const targetEnd = sourceXml.indexOf("</w:p>", targetStart) + "</w:p>".length;
      expect(savedXml?.startsWith(sourceXml.slice(0, targetStart))).toBe(true);
      expect(savedXml?.endsWith(sourceXml.slice(targetEnd))).toBe(true);
      expect(await saved.file("word/numbering.xml")?.async("text")).toBe(NUMBERING_XML);
    });
  }
});
