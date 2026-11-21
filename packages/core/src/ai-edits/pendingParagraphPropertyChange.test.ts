/**
 * A tracked paragraph-property edit to a paragraph whose properties already
 * carry a pending change. The canonical engine folds the edit into that one
 * change: it keeps the properties the change started from, so rejecting it
 * restores the original, and takes the latest edit's author and date. An
 * edit back to the original leaves nothing pending. The legacy engine
 * refuses the edit and leaves the pending change as it was.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

type BlockProperties = Extract<
  FolioDocumentOperation,
  { type: "setBlockParagraphProperties" }
>["properties"];

/** "Reviewer" centred the first paragraph, which stated no alignment. */
const source = async (): Promise<ArrayBuffer> => {
  const document: Document = fromMarkdown("Alpha text.\n\nTail.");
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") return panic("Expected a paragraph");
  paragraph.formatting = { alignment: "center" };
  paragraph.propertyChanges = [
    {
      type: "paragraphPropertyChange",
      info: { id: 7, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
      previousFormatting: {},
    },
  ];
  return await createDocx(document);
};

/** The first paragraph's `w:pPr` as saved, or "" when it states none. */
const savedProperties = async (bytes: ArrayBuffer): Promise<string> => {
  const xml =
    (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
    "";
  const paragraph = xml.match(/<w:p\b.*?<\/w:p>/su)?.[0] ?? "";
  const properties =
    paragraph.match(/^<w:p\b[^>]*>(<w:pPr>.*?<\/w:pPr>|<w:pPr\/>)/su)?.[1] ??
    "";
  return properties === "<w:pPr/>"
    ? ""
    : properties
        .replace(/ w:date="[^"]*"/gu, "")
        .replace(/ w:id="[^"]*"/gu, "");
};

type EditOptions = {
  session: "legacy" | "canonical";
  author: string;
  properties: BlockProperties;
};

const edit = async ({ session, author, properties }: EditOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
    session,
    author,
  });
  const block = reviewer.snapshot().blocks.at(0) ?? panic("Expected a block");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "edit",
          type: "setBlockParagraphProperties",
          blockId: block.id,
          properties,
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  const saved = await reviewer.toBuffer();
  const resolved = async (mode: "accept" | "reject") => {
    const copy = await FolioDocxReviewer.fromBuffer(saved.slice(0), {
      session,
    });
    if (mode === "accept") copy.acceptAll();
    else copy.rejectAll();
    return await savedProperties(await copy.toBuffer());
  };
  return {
    skipped: result.skipped.map(({ reason }) => reason),
    saved: await savedProperties(saved),
    accepted: await resolved("accept"),
    rejected: await resolved("reject"),
  };
};

const CHANGE = (author: string) =>
  `<w:pPrChange w:author="${author}"><w:pPr/></w:pPrChange>`;

describe("canonical: an edit over a pending paragraph-property change", () => {
  test("another author's edit folds into it, keeping where it started", async () => {
    const result = await edit({
      session: "canonical",
      author: "Editor",
      properties: { spacing: { spaceAfter: 360 } },
    });
    expect(result.skipped).toEqual([]);
    expect(result.saved).toBe(
      `<w:pPr><w:spacing w:after="360"/><w:jc w:val="center"/>${CHANGE("Editor")}</w:pPr>`,
    );
    expect(result.accepted).toBe(
      '<w:pPr><w:spacing w:after="360"/><w:jc w:val="center"/></w:pPr>',
    );
    expect(result.rejected).toBe("");
  });

  test("the same author's edit folds into it too", async () => {
    const result = await edit({
      session: "canonical",
      author: "Reviewer",
      properties: { alignment: "right" },
    });
    expect(result.saved).toBe(
      `<w:pPr><w:jc w:val="right"/>${CHANGE("Reviewer")}</w:pPr>`,
    );
    expect(result.rejected).toBe("");
  });

  test("an edit back to where it started leaves nothing pending", async () => {
    const result = await edit({
      session: "canonical",
      author: "Editor",
      properties: { alignment: null },
    });
    expect(result.skipped).toEqual([]);
    expect(result.saved).toBe("");
    expect(result.accepted).toBe("");
    expect(result.rejected).toBe("");
  });
});

describe("legacy: an edit over a pending paragraph-property change", () => {
  test("is refused, and the pending change stays as it was", async () => {
    const result = await edit({
      session: "legacy",
      author: "Editor",
      properties: { spacing: { spaceAfter: 360 } },
    });
    expect(result.skipped).toEqual(["pendingParagraphPropertyChange"]);
    expect(result.saved).toBe(
      `<w:pPr><w:jc w:val="center"/>${CHANGE("Reviewer")}</w:pPr>`,
    );
    expect(result.rejected).toBe("");
  });
});
