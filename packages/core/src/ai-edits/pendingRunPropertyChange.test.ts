/**
 * A tracked formatting edit to a run whose properties already carry a
 * pending change. The canonical engine folds the edit into that one change:
 * it keeps the properties the change started from, so rejecting it restores
 * the original, and takes the latest edit's author. An edit back to the
 * original leaves nothing pending.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";
import type { FolioAIInlineFormatting } from "./types";

/** "Reviewer" made the first word bold, which stated no properties. */
const source = async (): Promise<ArrayBuffer> => {
  const document: Document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000001",
      content: [
        {
          type: "run",
          formatting: { bold: true },
          propertyChanges: [
            {
              type: "runPropertyChange",
              info: { id: 7, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
              previousFormatting: {},
            },
          ],
          content: [{ type: "text", text: "Formatting" }],
        },
        { type: "run", content: [{ type: "text", text: " target" }] },
      ],
    },
  ];
  return createDocx(document);
};

/** The first run's `w:rPr` as saved, without dates and ids, or "" when it states none. */
const savedProperties = async (bytes: ArrayBuffer): Promise<string> => {
  const xml =
    (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
    "";
  const run = xml.match(/<w:r>.*?<\/w:r>/su)?.[0] ?? "";
  const properties =
    run.match(
      /^<w:r>(<w:rPr>.*?<\/w:rPr>(?!<\/w:rPrChange>)|<w:rPr\/>)/su,
    )?.[1] ?? "";
  return properties === "<w:rPr/>"
    ? ""
    : properties
        .replace(/ w:date="[^"]*"/gu, "")
        .replace(/ w:id="[^"]*"/gu, "")
        .replace(/ folio:[A-Za-z]+="[^"]*"/gu, "");
};

type EditOptions = {
  author: string;
  formatting: FolioAIInlineFormatting;
};

const edit = async ({ author, formatting }: EditOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
    author,
  });
  const block = reviewer.snapshot().blocks.at(0) ?? panic("Expected a block");
  const range =
    createFolioAITextRangeHandle({
      blockId: block.id,
      text: block.text,
      startOffset: 0,
      endOffset: "Formatting".length,
    }) ?? panic("Expected a range");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [{ id: "edit", type: "formatRange", range, formatting }],
    },
    { undefinedReferences: "refuse" },
  );
  const saved = await reviewer.toBuffer();
  const resolved = async (mode: "accept" | "reject") => {
    const copy = await FolioDocxReviewer.fromBuffer(saved.slice(0));
    if (mode === "accept") copy.acceptAll();
    else copy.rejectAll();
    return await savedProperties(await copy.toBuffer());
  };
  return {
    skipped: result.skipped.map(({ reason }) => reason),
    pending: reviewer
      .getChanges()
      .map((change) => ({ type: change.type, author: change.author })),
    saved: await savedProperties(saved),
    accepted: await resolved("accept"),
    rejected: await resolved("reject"),
  };
};

const CHANGE = (author: string) =>
  `<w:rPrChange w:author="${author}"><w:rPr/></w:rPrChange>`;

describe("canonical: an edit over a pending run-property change", () => {
  test("another author's edit folds into it, keeping where it started", async () => {
    const result = await edit({
      author: "Editor",
      formatting: { italic: true },
    });
    expect(result.skipped).toEqual([]);
    expect(result.pending).toEqual([{ type: "formatting", author: "Editor" }]);
    expect(result.saved).toBe(`<w:rPr><w:b/><w:i/>${CHANGE("Editor")}</w:rPr>`);
    expect(result.accepted).toBe("<w:rPr><w:b/><w:i/></w:rPr>");
    expect(result.rejected).toBe("");
  });

  test("the same author's edit folds into it too", async () => {
    const result = await edit({
      author: "Reviewer",
      formatting: { italic: true },
    });
    expect(result.saved).toBe(
      `<w:rPr><w:b/><w:i/>${CHANGE("Reviewer")}</w:rPr>`,
    );
    expect(result.rejected).toBe("");
  });

  test("an edit back to where it started leaves nothing pending", async () => {
    const result = await edit({
      author: "Editor",
      formatting: { bold: null },
    });
    expect(result.skipped).toEqual([]);
    expect(result.pending).toEqual([]);
    expect(result.saved).toBe("");
    expect(result.accepted).toBe("");
    expect(result.rejected).toBe("");
  });
});
