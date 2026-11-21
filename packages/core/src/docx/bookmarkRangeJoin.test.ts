import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";

const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const packageWith = async (paragraph: string): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}"><w:body>${paragraph}<w:p><w:r><w:t>Tail</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  );
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer as ArrayBuffer;
};

/** The bookmark markers of the saved body, in document order. */
type MarkersAfterSplitOptions = {
  paragraph: string;
  mode: "direct" | "tracked-changes";
  session: HeadlessReviewerSession;
};

const markersAfterSplit = async ({
  paragraph,
  mode,
  session,
}: MarkersAfterSplitOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await packageWith(paragraph),
    { session },
  );
  const blockId = reviewer.snapshot().blocks[0]?.id;
  if (blockId === undefined) throw new Error("the fixture has no paragraph");
  const result = reviewer.applyDocumentOperations(
    {
      version: 1,
      mode,
      operations: [{ id: "split", type: "splitBlock", blockId, offset: 6 }],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.applied).toHaveLength(1);
  const saved = await ZipArchive.read(await reviewer.toBuffer());
  const xml = (await saved.entry("word/document.xml")?.text()) ?? "";
  return {
    markers: [
      ...xml.matchAll(/<w:bookmark(Start|End) [^>]*w:id="(\d+)"/gu),
    ].map(([, kind, id]) => `${kind}:${id}`),
    text: [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)</gu)]
      .map(([, text]) => text)
      .join(""),
  };
};

describe("a paragraph split keeps one range per bookmark", () => {
  for (const session of SESSIONS)
    for (const mode of ["direct", "tracked-changes"] as const) {
      // Legacy closes a start the source left open; the canonical session
      // writes the start as the source has it. Neither writes it twice.
      test(`${session}: a start with no end in the source is written once, ${mode}`, async () => {
        const { markers, text } = await markersAfterSplit({
          paragraph:
            '<w:p><w:r><w:t>Alpha beta gamma</w:t></w:r><w:bookmarkStart w:id="1" w:name="_GoBack"/></w:p>',
          mode,
          session,
        });
        expect(markers).toEqual(
          session === "legacy" ? ["Start:1", "End:1"] : ["Start:1"],
        );
        expect(text).toBe("Alpha beta gammaTail");
      });

      test(`${session}: a range inside the paragraph stays as it was, ${mode}`, async () => {
        const { markers } = await markersAfterSplit({
          paragraph:
            '<w:p><w:r><w:t xml:space="preserve">Alpha beta </w:t></w:r><w:bookmarkStart w:id="1" w:name="mark"/><w:r><w:t>gamma</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>',
          mode,
          session,
        });
        expect(markers).toEqual(["Start:1", "End:1"]);
      });
    }
});
