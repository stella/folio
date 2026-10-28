/**
 * What a session writes for paragraphs it did not author.
 *
 * - An inserted paragraph copies its anchor's formatting but not the
 *   attributes of the `w:p` it is not (`w:rsidR`), which the save never
 *   writes for it.
 * - A deletion read from a package is reconciled against the baseline it was
 *   read with, so a later save does not turn the language it inherited into
 *   direct formatting.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  );
  zip.set(
    "word/styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:lang w:val="fr-FR" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
  );
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const open = async (bytes: ArrayBuffer): Promise<FolioDocxReviewer> =>
  await FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });

const documentXml = async (bytes: ArrayBuffer): Promise<string> => {
  const xml = await (
    await ZipArchive.read(bytes)
  )
    .entry("word/document.xml")
    ?.text();
  if (xml === undefined) throw new Error("no document part");
  return xml;
};

const insertAfter = (reviewer: FolioDocxReviewer, text: string) => {
  const anchor = reviewer.getContent().at(-1);
  if (!anchor) throw new Error("no anchor");
  return reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        { id: "insert", type: "insertAfterBlock", blockId: anchor.id, text },
      ],
    },
    { undefinedReferences: "refuse" },
  );
};

describe("an inserted paragraph", () => {
  test("copies its anchor's formatting but not its source attributes", async () => {
    const reviewer = await open(
      await buildDocx(
        `<w:p w:rsidR="006B7EF7" w:rsidRDefault="00A048B5"><w:r><w:t>Anchor.</w:t></w:r></w:p>`,
      ),
    );
    expect(insertAfter(reviewer, "Added.").applied).toHaveLength(1);
    const [anchor, added] = reviewer.toDocument().package.document.content;
    expect(anchor).toMatchObject({
      preservedAttributes: [
        { name: "rsidR", value: "006B7EF7" },
        { name: "rsidRDefault", value: "00A048B5" },
      ],
    });
    expect(added).toBeDefined();
    expect(added).not.toHaveProperty("preservedAttributes");
    // The live paragraph is what accepting the insertion keeps.
    const paragraphs: unknown[] = [];
    reviewer.state.doc.forEach((node) => {
      paragraphs.push(node.attrs["_preservedAttributes"] ?? null);
    });
    expect(paragraphs.at(0)).not.toBeNull();
    expect(paragraphs.at(1)).toBeNull();
  });
});

describe("a deletion read from a package", () => {
  test("keeps the language it inherited out of the next save", async () => {
    const deletion =
      `<w:p><w:pPr><w:rPr><w:i/><w:lang w:val="en-US"/></w:rPr></w:pPr>` +
      `<w:del w:id="2" w:author="Reviewer" w:date="2026-10-19T00:00:00Z">` +
      `<w:r><w:rPr><w:i/><w:sz w:val="24"/></w:rPr><w:delText>Gone</w:delText></w:r>` +
      `</w:del></w:p>`;
    const reviewer = await open(
      await buildDocx(`${deletion}<w:p><w:r><w:t>Kept.</w:t></w:r></w:p>`),
    );
    expect(insertAfter(reviewer, "Added.").applied).toHaveLength(1);
    const saved = await documentXml(await reviewer.toBuffer());
    const run = /<w:del\b[^>]*>(.*?)<\/w:del>/su.exec(saved)?.[1] ?? "";
    expect(run).toContain("Gone");
    expect(run).not.toContain("w:lang");
  });
});
