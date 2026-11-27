import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const KEPT_ID = "1A2B3C4D";

/** A package as many producers write it: no paragraph ids but one. */
const paraIdLessDocx = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  const options = { modified: new Date("2026-01-01T00:00:00Z") };
  zip.set(
    "[Content_Types].xml",
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    options,
  );
  zip.set(
    "_rels/.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    options,
  );
  zip.set(
    "word/document.xml",
    `<w:document xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>
<w:p><w:r><w:t>First</w:t></w:r></w:p>
<w:p w14:paraId="${KEPT_ID}"><w:r><w:t>Kept</w:t></w:r></w:p>
<w:p><w:r><w:t>Same</w:t></w:r></w:p>
<w:p><w:r><w:t>Same</w:t></w:r></w:p>
<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:p/>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
</w:body></w:document>`,
    options,
  );
  return (await zip.generate({ level: 6 })).buffer as ArrayBuffer;
};

const documentXml = async (docx: ArrayBuffer): Promise<string> => {
  const zip = await ZipArchive.read(docx);
  const text = await zip.entry("word/document.xml")?.text();
  if (text === undefined) throw new Error("the package has no main document");
  return text;
};

/** Each paragraph's `[paraId, textId]`, in document order. */
const paragraphIds = (
  xml: string,
): [string | undefined, string | undefined][] =>
  [...xml.matchAll(/<w:p(?=[\s/>])[^>]*>/gu)].map(([tag]) => [
    /\bw14:paraId="([^"]*)"/u.exec(tag)?.[1],
    /\bw14:textId="([^"]*)"/u.exec(tag)?.[1],
  ]);

const save = async (docx: ArrayBuffer): Promise<ArrayBuffer> =>
  repackDocx(await parseDocx(docx, { preloadFonts: false }), {
    updateModifiedDate: false,
  });

describe("a save gives every paragraph a paragraph id", () => {
  test("each paragraph holds a valid id, unique in the package, with a text id beside a new one", async () => {
    const ids = paragraphIds(
      await documentXml(await save(await paraIdLessDocx())),
    );
    expect(ids).toHaveLength(6);
    const paraIds = ids.map(([paraId]) => paraId);
    for (const paraId of paraIds) {
      expect(paraId).toMatch(/^[0-9A-F]{8}$/u);
      expect(Number.parseInt(paraId ?? "", 16)).toBeGreaterThan(0);
      expect(Number.parseInt(paraId ?? "", 16)).toBeLessThan(0x80000000);
    }
    expect(new Set(paraIds).size).toBe(paraIds.length);
    // The id the producer wrote stays; the minted ones carry a text id.
    expect(paraIds[1]).toBe(KEPT_ID);
    for (const [index, [paraId, textId]] of ids.entries())
      if (index !== 1) expect(textId).toBe(paraId);
  });

  test("a second save writes the same ids", async () => {
    const first = await save(await paraIdLessDocx());
    const second = await save(first);
    expect(paragraphIds(await documentXml(second))).toEqual(
      paragraphIds(await documentXml(first)),
    );
    // Saving the same paraId-less package twice mints the same ids.
    expect(
      paragraphIds(await documentXml(await save(await paraIdLessDocx()))),
    ).toEqual(paragraphIds(await documentXml(first)));
  });
});
