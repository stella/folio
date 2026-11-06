/**
 * An inline content control cannot span paragraphs, so a split inside one is
 * refused, in both modes, and leaves the document as it was. A split beside the
 * control still applies.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const open = async (): Promise<FolioDocxReviewer> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:p w14:paraId="10000001"><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="7"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Inside</w:t></w:r></w:sdtContent></w:sdt></w:p><w:p w14:paraId="10000002"><w:r><w:t>Next</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  );
  return FolioDocxReviewer.fromBuffer((await zip.generate()).slice().buffer);
};

const split = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked-changes",
  offset: number,
) =>
  reviewer.applyDocumentOperations(
    {
      version: 1,
      mode,
      operations: [
        { id: "split", type: "splitBlock", blockId: "10000001", offset },
      ],
    } as never,
    { undefinedReferences: "refuse" },
  );

const texts = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ text }) => text);

describe.each(["direct", "tracked-changes"] as const)("%s", (mode) => {
  test("a split inside the control is refused and changes nothing", async () => {
    const reviewer = await open();
    const before = texts(reviewer);
    const result = split(reviewer, mode, "Before In".length);
    expect(result.applied).toEqual([]);
    expect(texts(reviewer)).toEqual(before);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
    );
    reopened.rejectAll();
    expect(texts(reopened)).toEqual(before);
  });

  test("a split beside the control applies", async () => {
    const reviewer = await open();
    const result = split(reviewer, mode, "Bef".length);
    expect(result.applied).toHaveLength(1);
    expect(texts(reviewer).slice(0, 2)).toEqual(["Bef", "ore Inside"]);
  });
});
