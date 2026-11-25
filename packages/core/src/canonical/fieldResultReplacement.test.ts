import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { buildBodySequenceDocx } from "../compare/__fixtures__/body-sequence";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const REFERENCE =
  '<w:fldSimple w:instr=" REF _Ref1 \\h "><w:r><w:t>clause 1</w:t></w:r></w:fldSimple>';

const documentPartOf = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await ZipArchive.read(buffer)).entry("word/document.xml")?.text()) ??
  "";

/** One paragraph: "See clause 1 now." with "clause 1" a field's result. */
const docxWithReference = async (): Promise<ArrayBuffer> => {
  const buffer = await buildBodySequenceDocx([
    { kind: "paragraph", text: "Placeholder." },
  ]);
  const zip = await ZipArchive.read(buffer);
  const xml = await documentPartOf(buffer);
  const body = xml.indexOf("<w:body>");
  const section = xml.indexOf("<w:sectPr", body);
  if (body === -1 || section === -1) {
    throw new Error("The fixture body has no section");
  }
  const open = body + "<w:body>".length;
  zip.set(
    "word/document.xml",
    xml.slice(0, open) +
      '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>' +
      REFERENCE +
      '<w:r><w:t xml:space="preserve"> now.</w:t></w:r></w:p>' +
      xml.slice(section),
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const firstParagraphText = (xml: string): string =>
  [
    ...(/<w:p[ >][\s\S]*?<\/w:p>/u.exec(xml)?.[0] ?? "").matchAll(
      /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/gu,
    ),
  ]
    .map(([, text]) => text)
    .join("");

describe.each(SESSIONS)("%s session", (session) => {
  // The word diff puts the new word inside the field's result; the field is
  // replaced whole by the text, so accepting leaves the edited words and
  // rejecting gives the field back.
  test.each(["direct", "tracked-changes"] as const)(
    "a %s replacement that edits a field's result replaces the field with its text",
    async (mode) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await docxWithReference(),
        { session },
      );
      const blockId = reviewer.snapshot().blocks.at(0)?.id ?? "";
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            {
              id: "edit",
              type: "replaceBlock",
              blockId,
              text: "See clause new 1 now.",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([]);
      expect(result.applied.map(({ id }) => id)).toEqual(["edit"]);

      const edited = await reviewer.toBuffer();
      const accepted = await FolioDocxReviewer.fromBuffer(edited, { session });
      accepted.acceptAll();
      const acceptedXml = await documentPartOf(await accepted.toBuffer());
      expect(firstParagraphText(acceptedXml)).toBe("See clause new 1 now.");
      expect(acceptedXml).not.toContain("REF _Ref1");
      if (mode === "direct") return;

      const rejected = await FolioDocxReviewer.fromBuffer(edited, { session });
      rejected.rejectAll();
      const rejectedXml = await documentPartOf(await rejected.toBuffer());
      expect(firstParagraphText(rejectedXml)).toBe("See clause 1 now.");
      expect(rejectedXml).toContain("REF _Ref1");
    },
  );
});
