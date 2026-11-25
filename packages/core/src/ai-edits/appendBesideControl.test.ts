/**
 * Text a replacement adds at an inline content control's edge belongs to the
 * control only when the replaced text is the control's own: a replacement
 * reaching past the control adds it beside the control.
 */
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const CONTROL = `<w:sdt><w:sdtPr><w:alias w:val="Name"/><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:r><w:t>John</w:t></w:r></w:sdtContent></w:sdt>`;

const open = async (session: HeadlessReviewerSession) => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:p w14:paraId="10000001"><w:r><w:t xml:space="preserve">Name: </w:t></w:r>${CONTROL}</w:p></w:body></w:document>`,
  );
  return FolioDocxReviewer.fromBuffer((await zip.generate()).slice().buffer, {
    session,
  });
};

const controlText = async (reviewer: FolioDocxReviewer): Promise<string> => {
  const zip = await ZipArchive.read(await reviewer.toBuffer());
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  const content = /<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/u.exec(xml)?.[1];
  return [...(content ?? "").matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/gu)]
    .map(([, text]) => text)
    .join("");
};

describe.each([
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[])(
  "text added at a content control's end (%s)",
  (session) => {
    test.each([
      ["the paragraph", "Name: John", "Name: John!", "John"],
      ["the control's text", "John", "John!", "John!"],
    ])(
      "replacing %s keeps the control's text as the range owns it",
      async (_name, find, replace, expected) => {
        const reviewer = await open(session);
        const result = reviewer.applyOperations(
          [
            {
              id: "edit",
              type: "replaceInBlock",
              blockId: "10000001",
              find,
              replace,
            },
          ],
          { undefinedReferences: "refuse", mode: "direct" },
        );
        expect(result.skipped).toEqual([]);
        expect(reviewer.getContent().at(0)?.text).toBe("Name: John!");
        expect(await controlText(reviewer)).toBe(expected);
      },
    );
  },
);
