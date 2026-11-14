import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { parseDocx } from "./parser";
import type { SelectiveSaveDecline } from "./saveDiagnostics";
import { attemptSelectiveSave } from "./selectiveSave";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

/** Two identified paragraphs; the second may hold a bookmark. */
const source = async (bookmark: boolean): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  const marked = bookmark
    ? `<w:bookmarkStart w:id="0" w:name="Mark"/><w:r><w:t>Marked</w:t></w:r><w:bookmarkEnd w:id="0"/>`
    : `<w:r><w:t>Marked</w:t></w:r>`;
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body><w:p w14:paraId="11111111" w14:textId="11111111"><w:r><w:t>Plain</w:t></w:r></w:p><w:p w14:paraId="22222222" w14:textId="22222222">${marked}</w:p></w:body></w:document>`,
  );
  return (await zip.generate()).slice().buffer;
};

/** Insert an empty paragraph and save selectively, collecting declines. */
const insertAndSave = async (bookmark: boolean, maxBytes?: number) => {
  const bytes = await source(bookmark);
  const document = await parseDocx(bytes, {
    preloadFonts: false,
    detectVariables: false,
  });
  document.package.document.content.splice(1, 0, {
    type: "paragraph",
    paraId: "33333333",
    content: [],
  });
  const declines: SelectiveSaveDecline[] = [];
  const saved = await attemptSelectiveSave(document, bytes, {
    changedParaIds: new Set(["33333333"]),
    structuralChange: true,
    hasUntrackedChanges: false,
    ...(maxBytes === undefined ? {} : { maxBytes }),
    onSelectiveDecline: (decline) => declines.push(decline),
  });
  return { saved, declines };
};

describe("selective save declines", () => {
  test("a bookmark in the body declines a structural splice by name", async () => {
    const { saved, declines } = await insertAndSave(true);
    expect(saved).toBeNull();
    expect(declines).toEqual([
      { part: "word/document.xml", reason: "structuralRangeElement" },
    ]);
  });

  test("a body without range elements is spliced with no decline", async () => {
    const { saved, declines } = await insertAndSave(false);
    expect(saved).not.toBeNull();
    expect(declines).toEqual([]);
  });

  test("a source over the size limit declines as too large", async () => {
    const { saved, declines } = await insertAndSave(false, 1);
    expect(saved).toBeNull();
    expect(declines).toEqual([
      { part: "word/document.xml", reason: "tooLarge" },
    ]);
  });
});
