/**
 * Operations on a paragraph whose mark is pending deletion: the canonical
 * reviewer applies what the default one applies there and refuses what it
 * refuses, with the same reason, directly and as a tracked change, and the
 * two read the same words pending, accepted and rejected, also after save and
 * reopen; a paragraph inserted after such a paragraph, where the canonical
 * reviewer puts it right after it, reads as each case states. A test-only break that has the canonical owner refuse every such
 * edit shows the law fails.
 */

import { describe, expect, mock, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import type { FolioDocumentOperation } from "../document-operations";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";
import * as pendingJoin from "./pendingJoin";
import { hashFolioAIBlockText } from "./snapshot";
import type { FolioAIEditApplyMode } from "./types";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
/** A paragraph mark pending deletion, as revision `id`. */
const deletedMark = (id: number): string =>
  `<w:pPr><w:rPr><w:del w:id="${String(id)}" w:author="Author" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>`;

const paragraph = (text: string, properties = ""): string =>
  `<w:p>${properties}${text === "" ? "" : `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`}</w:p>`;

/** A joined pair, a blank paragraph that goes whole, and a closing one. */
const BODY =
  paragraph("First words here", deletedMark(1)) +
  paragraph("Second words") +
  paragraph("", deletedMark(2)) +
  paragraph("Closing words");

const fixture = async (): Promise<ArrayBuffer> => {
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
    `${HEAD}<w:document xmlns:w="${W}"><w:body>${BODY}<w:sectPr/></w:body></w:document>`,
  );
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
};

type Case = {
  name: string;
  operation: (ids: { joined: string; blank: string }) => FolioDocumentOperation;
  /**
   * What the canonical reviewer reads where it departs from the default one:
   * a paragraph inserted after one with words whose mark is pending deletion
   * goes right after it, as pressing Enter at its end does, and the join then
   * runs into it; the default engine puts it after the paragraph the join
   * leaves (`insertionAfterPendingBreak`).
   */
  canonical?: (mode: FolioAIEditApplyMode) => unknown;
};

/** The words a case reads pending, reopened, accepted and rejected. */
const insertedRightAfter = (
  mode: FolioAIEditApplyMode,
  pending: string[],
  accepted: string[],
  inserted: string,
) => ({
  outcome: { applied: 1, skipped: [] },
  pending,
  reopened: pending,
  accepted,
  rejected:
    mode === "direct" ? pending : pending.filter((text) => text !== inserted),
});

const CASES: readonly Case[] = [
  {
    name: "a replacement inside it",
    operation: ({ joined }) => ({
      id: "op",
      type: "replaceInBlock",
      blockId: joined,
      find: "words",
      replace: "terms",
    }),
  },
  {
    name: "a block after it",
    operation: ({ joined }) => ({
      id: "op",
      type: "insertAfterBlock",
      blockId: joined,
      text: "Inserted after",
    }),
    canonical: (mode) =>
      insertedRightAfter(
        mode,
        ["First words here", "Inserted after", "Second words", "Closing words"],
        ["First words hereInserted after", "Second words", "Closing words"],
        "Inserted after",
      ),
  },
  {
    name: "a block before it",
    operation: ({ joined }) => ({
      id: "op",
      type: "insertBeforeBlock",
      blockId: joined,
      text: "Inserted before",
    }),
  },
  {
    name: "formatting a range of it",
    operation: ({ joined }) => ({
      id: "op",
      type: "formatRange",
      range: {
        type: "textRange",
        story: "main",
        blockId: joined,
        startOffset: 0,
        endOffset: 5,
        selectedTextHash: hashFolioAIBlockText("First"),
      },
      formatting: { bold: true },
    }),
  },
  {
    name: "its paragraph properties",
    operation: ({ joined }) => ({
      id: "op",
      type: "setBlockParagraphProperties",
      blockId: joined,
      properties: { alignment: "center" },
    }),
  },
  {
    name: "replacing its whole text",
    operation: ({ joined }) => ({
      id: "op",
      type: "replaceBlock",
      blockId: joined,
      text: "Other words here",
    }),
  },
  {
    name: "replacing a blank paragraph that goes whole",
    operation: ({ blank }) => ({
      id: "op",
      type: "replaceBlock",
      blockId: blank,
      text: "Written there",
    }),
  },
  {
    name: "a block after a blank paragraph that goes whole",
    operation: ({ blank }) => ({
      id: "op",
      type: "insertAfterBlock",
      blockId: blank,
      text: "Inserted after blank",
    }),
  },
];

const MODES: readonly FolioAIEditApplyMode[] = ["direct", "tracked-changes"];

const open = (buffer: ArrayBuffer, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Reviewer",
    session,
  });

const save = (reviewer: FolioDocxReviewer) => reviewer.toBuffer();

/** The words a saved body shows and the words it holds as deleted. */
const savedWords = async (buffer: ArrayBuffer) => {
  const zip = await ZipArchive.read(buffer);
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  const paragraphs = xml
    .split(/<\/w:p>/u)
    .map((part) =>
      [...part.matchAll(/<w:t(?: [^>]*)?>([^<]*)</gu)]
        .map(([, text]) => text)
        .join(""),
    );
  return paragraphs.filter((text) => text !== "");
};

/** The same block ids on both engines: by the text each block reads. */
const idsOf = (reviewer: FolioDocxReviewer) => {
  const blocks = reviewer.snapshot().blocks;
  const joined = blocks.find((block) => block.text === "First words here");
  const blank = blocks.at(blocks.findIndex((block) => block === joined) + 2);
  if (joined === undefined || blank === undefined)
    throw new Error("the fixture's blocks are missing");
  return { joined: joined.id, blank: blank.id };
};

const runCase = async (
  buffer: ArrayBuffer,
  session: HeadlessReviewerSession,
  { operation }: Case,
  mode: FolioAIEditApplyMode,
) => {
  const apply = async () => {
    const reviewer = await open(buffer, session);
    const result = reviewer.applyDocumentOperations(
      { version: 1, mode, operations: [operation(idsOf(reviewer))] },
      { undefinedReferences: "refuse" },
    );
    return { reviewer, result };
  };
  const { reviewer, result } = await apply();
  const outcome = {
    applied: result.applied.length,
    skipped: result.skipped.map(({ reason }) => reason),
  };
  if (outcome.applied === 0) return { outcome };
  const pendingSaved = await save(reviewer);
  const accepted = (await apply()).reviewer;
  accepted.acceptAll();
  const rejected = (await apply()).reviewer;
  rejected.rejectAll();
  const reopened = await open(pendingSaved, session);
  return {
    outcome,
    pending: await savedWords(pendingSaved),
    reopened: await savedWords(await save(reopened)),
    accepted: await savedWords(await save(accepted)),
    rejected: await savedWords(await save(rejected)),
  };
};

describe("an operation on a paragraph whose mark is pending deletion", () => {
  for (const testCase of CASES)
    for (const mode of MODES)
      test(`${mode}: ${testCase.name}`, async () => {
        const buffer = await fixture();
        const canonical = await runCase(buffer, "canonical", testCase, mode);
        if (testCase.canonical !== undefined) {
          expect(canonical).toEqual(testCase.canonical(mode));
          return;
        }
        const legacy = await runCase(buffer, "legacy", testCase, mode);
        expect(canonical).toEqual(legacy);
      }, 60_000);
});

test("the law catches a canonical session that refuses every edit on such a paragraph", async () => {
  // The test-only break: the canonical compiler's owner answers with the
  // generic refusal it gave before for every operation that places no block,
  // while the default engine keeps its own rules.
  const original = { ...pendingJoin };
  mock.module("./pendingJoin", () => ({
    ...original,
    pendingMarkEditRefusal: (
      options: Parameters<typeof original.pendingMarkEditRefusal>[0],
    ) =>
      original.isDeletedPPrMark(options.paragraph.attrs["pPrMark"])
        ? original.PENDING_WHOLE_PARAGRAPH_REFUSAL
        : original.pendingMarkEditRefusal(options),
  }));
  try {
    const buffer = await fixture();
    let diverging = 0;
    for (const testCase of CASES.filter(
      ({ canonical }) => canonical === undefined,
    ))
      for (const mode of MODES) {
        const legacy = await runCase(buffer, "legacy", testCase, mode);
        const canonical = await runCase(buffer, "canonical", testCase, mode);
        if (!Bun.deepEquals(canonical, legacy)) diverging += 1;
      }
    expect(diverging).toBeGreaterThan(0);
  } finally {
    mock.module("./pendingJoin", () => original);
  }
}, 120_000);
