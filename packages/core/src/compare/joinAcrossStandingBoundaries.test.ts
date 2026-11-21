import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { buildBodySequenceDocx } from "./__fixtures__/body-sequence";
import { compareDocx } from "./compare";

const SESSIONS = [
  "canonical",
  "legacy",
] as const satisfies readonly HeadlessReviewerSession[];

const OPTIONS = {
  author: "compare",
  timestamp: "2026-11-21T00:00:00.000Z",
} as const;

const PAGE_FIELD =
  '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>';

const paragraph = (text: string, tail = ""): string =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r>${tail}</w:p>`;

const bookmarkStart = (id: number): string =>
  `<w:bookmarkStart w:id="${String(id)}" w:name="_Ref${String(id)}"/>`;

const bookmarkEnd = (id: number): string =>
  `<w:bookmarkEnd w:id="${String(id)}"/>`;

const documentPartOf = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await ZipArchive.read(buffer)).entry("word/document.xml")?.text()) ??
  "";

/** A package whose body holds `blocks` before the fixture's section. */
const docxWithBody = async (blocks: string): Promise<ArrayBuffer> => {
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
    xml.slice(0, open) + blocks + xml.slice(section),
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const paragraphTexts = (xml: string): string[] =>
  [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/gu)].map(([block]) =>
    [...block.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/gu)]
      .map(([, text]) => text)
      .join(""),
  );

describe.each(SESSIONS)("%s session", (session) => {
  // The paragraphs the revised document joins have bookmark boundaries
  // standing between them; a field elsewhere in the story makes the
  // comparison line every paragraph's inline atoms up with the target's.
  test("joins two paragraphs across standing bookmarks beside a field", async () => {
    const base = await docxWithBody(
      bookmarkStart(1) +
        paragraph("First words.") +
        bookmarkEnd(1) +
        bookmarkStart(2) +
        paragraph("Second words.") +
        bookmarkEnd(2) +
        paragraph("Page ", PAGE_FIELD),
    );
    const target = await docxWithBody(
      bookmarkStart(1) +
        bookmarkEnd(1) +
        bookmarkStart(2) +
        paragraph("First words. Second words.") +
        bookmarkEnd(2) +
        paragraph("Page ", PAGE_FIELD),
    );

    const compared = await compareDocx(base, target, { ...OPTIONS, session });
    if (compared.isErr()) throw compared.error;
    expect(compared.value.verification).toEqual({ status: "verified" });

    for (const [view, expected] of [
      ["acceptAll", ["First words. Second words.", "Page 1"]],
      ["rejectAll", ["First words.", "Second words.", "Page 1"]],
    ] as const) {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        compared.value.buffer,
        { session },
      );
      expect(reviewer[view]().skipped).toEqual([]);
      const xml = await documentPartOf(await reviewer.toBuffer());
      expect(paragraphTexts(xml)).toEqual([...expected]);
      expect(xml).toContain('w:instr=" PAGE "');
    }
  });
});
