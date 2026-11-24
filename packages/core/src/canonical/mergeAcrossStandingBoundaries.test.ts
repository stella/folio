/**
 * Merging a paragraph with the next one when bookmark boundaries stand
 * between them in the body: both engines merge across them. Directly, the
 * boundaries end up where the two paragraphs meet; tracked, accepting gives
 * the direct result and rejecting gives the source back.
 */
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createEmptyDocx } from "../docx/rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

const packageOf = async (body: string): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const paragraph = (paraId: string, text: string) =>
  `<w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p>`;

const SHAPES = {
  "a bookmark end": `<w:bookmarkStart w:id="1" w:name="Clause"/>${paragraph("10000001", "Alpha")}<w:bookmarkEnd w:id="1"/>${paragraph("10000002", "Beta")}`,
  "a bookmark start": `${paragraph("10000001", "Alpha")}<w:bookmarkStart w:id="2" w:name="Later"/>${paragraph("10000002", "Beta")}<w:bookmarkEnd w:id="2"/>`,
} as const;

/** What a reader sees: the words, and which bookmark ranges hold them. */
const reading = async (reviewer: FolioDocxReviewer) => {
  const xml =
    (await (
      await ZipArchive.read(await reviewer.toBuffer())
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return {
    blocks: reviewer.snapshot().blocks.map(({ text }) => text),
    marks: [
      ...xml.matchAll(/<w:(bookmark(?:Start|End))[^>]*w:id="(\d+)"/gu),
    ].map(([, kind, id]) => `${kind}:${id}`),
    text: [...xml.matchAll(/<w:t>([^<]*)<\/w:t>|<w:bookmark(Start|End)\b/gu)]
      .map(([, text, boundary]) => text ?? boundary)
      .join("|"),
  };
};

describe.each(["canonical"] as const)("%s", (session) => {
  for (const [shape, body] of Object.entries(SHAPES))
    test(`merges across ${shape}`, async () => {
      const bytes = await packageOf(body);
      const open = (source: ArrayBuffer) =>
        FolioDocxReviewer.fromBuffer(source, { session });
      const run = async (mode: "direct" | "tracked-changes") => {
        const reviewer = await open(bytes);
        const result = reviewer.applyDocumentOperations(
          {
            version: 1,
            mode,
            operations: [
              { id: "merge", type: "mergeBlockWithNext", blockId: "10000001" },
            ],
          },
          { undefinedReferences: "refuse" },
        );
        expect(result.skipped).toEqual([]);
        return reviewer;
      };
      const direct = await reading(await run("direct"));
      expect(direct.blocks).toEqual(["AlphaBeta"]);
      const tracked = await (await run("tracked-changes")).toBuffer();
      const accepted = await open(tracked);
      accepted.acceptAll();
      expect(await reading(accepted)).toEqual(direct);
      const rejected = await open(tracked);
      rejected.rejectAll();
      expect(await reading(rejected)).toEqual(await reading(await open(bytes)));
    });
});
