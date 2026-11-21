/**
 * A replacement whose text repeats a tab the paragraph holds keeps the tab
 * and changes only the words around it, on either session.
 */
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const open = async (session: HeadlessReviewerSession) => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:p w14:paraId="10000001"><w:r><w:t>Term</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>12 months</w:t></w:r></w:p></w:body></w:document>`,
  );
  return FolioDocxReviewer.fromBuffer((await zip.generate()).slice().buffer, {
    session,
  });
};

const replace = (
  reviewer: FolioDocxReviewer,
  find: string,
  replacement: string,
) =>
  reviewer.applyOperations(
    [
      {
        id: "edit",
        type: "replaceInBlock",
        blockId: "10000001",
        find,
        replace: replacement,
      },
    ],
    { undefinedReferences: "refuse", mode: "direct" },
  );

describe.each([
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[])(
  "a replacement across a tab (%s)",
  (session) => {
    test("keeps the tab it repeats", async () => {
      const reviewer = await open(session);
      const result = replace(reviewer, "Term\t12 months", "Term\t24 months");
      expect(result.skipped).toEqual([]);
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        { session },
      );
      expect(reopened.getContent().at(0)?.text).toBe("Term\t24 months");
    });
  },
);
