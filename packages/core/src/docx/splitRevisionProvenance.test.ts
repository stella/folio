/**
 * A tracked change holding a hyperlink is written as one element per stretch,
 * the revision inside the link. Its resolution joins describe the whole
 * change, so they travel as split pieces that a reader puts back together.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { Document, ParagraphContent } from "../types/document";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const JOINS = { before: 1, after: 2, remove: 1 };

const documentOf = (change: ParagraphContent): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "10000000",
          content: [run("a"), change, run("e")],
        },
      ],
    },
  },
});

const changeOf = (type: "insertion" | "deletion"): ParagraphContent => ({
  type,
  info: { id: 7, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
  resolutionJoins: JOINS,
  content: [
    run("b"),
    { type: "hyperlink", href: "https://example.org/", children: [run("c")] },
    run("d"),
  ],
});

const documentXml = async (bytes: Uint8Array | ArrayBuffer) =>
  (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
  "";

const paragraphContent = (document: Document) => {
  const [paragraph] = document.package.document.content;
  return paragraph?.type === "paragraph" ? paragraph.content : [];
};

describe("a tracked change around a hyperlink", () => {
  for (const type of ["insertion", "deletion"] as const)
    test(`a ${type} saves as pieces and reopens as the one change with its joins`, async () => {
      const bytes = await createDocx(documentOf(changeOf(type)));
      const xml = await documentXml(bytes);
      // Three pieces, the middle one inside the link; no piece claims the
      // whole change's joins under the attribute older readers parse.
      expect(xml.match(/folio:splitResolutionJoins=/gu)?.length).toBe(3);
      expect(xml).not.toContain("folio:resolutionJoins=");
      expect(xml).toMatch(/mc:Ignorable="[^"]*\bfolio\b/u);
      const reopened = await parseDocx(bytes, { preloadFonts: false });
      const content = paragraphContent(reopened);
      expect(content.map(({ type: kind }) => kind)).toEqual([
        "run",
        type,
        "run",
      ]);
      const change = content[1];
      expect(
        change?.type === type && {
          id: change.info.id,
          joins: change.resolutionJoins,
          kinds: change.content.map(({ type: kind }) => kind),
        },
      ).toEqual({
        id: 7,
        joins: JOINS,
        kinds: ["run", "hyperlink", "run"],
      });
    });

  test("a reader that ignores the pieces keeps the revisions, which resolve", async () => {
    const bytes = await createDocx(documentOf(changeOf("insertion")));
    // What a reader that does not know the attribute sees.
    const zip = await ZipArchive.read(bytes);
    zip.set(
      "word/document.xml",
      (await documentXml(bytes)).replaceAll(
        / folio:splitResolutionJoins="[^"]*"/gu,
        "",
      ),
    );
    const stripped = await zip.generate({ compression: "store" });
    const reopened = await parseDocx(stripped.slice().buffer, {
      preloadFonts: false,
    });
    expect(
      paragraphContent(reopened).filter(({ type }) => type === "insertion"),
    ).toHaveLength(3);
    const reviewer = await FolioDocxReviewer.fromBuffer(
      stripped.slice().buffer,
    );
    reviewer.rejectAll();
    expect(reviewer.getContent().map(({ text }) => text)).toEqual(["ae"]);
  });
});
