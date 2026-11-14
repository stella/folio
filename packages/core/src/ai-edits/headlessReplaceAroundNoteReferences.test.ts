/**
 * A replacement over a range that holds note references rewrites the prose
 * around them and keeps each reference where it stands, on both reviewer
 * engines, directly and as a tracked change: the engines apply or refuse the
 * same operations with the same reason, read the same text pending, accepted
 * and rejected, and save every reference, also after reopening.
 */

import { describe, expect, mock, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { fromMarkdown } from "../markdown";
import { createDocx, ensureParaIds } from "../server";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Endnote, Footnote, Paragraph } from "../types/document";
import { canonicalReviewerSave } from "./__tests__/canonicalReviewerSave";
import { FolioDocxReviewer } from "./headless";
import * as noteReferences from "./note-references";
import type { FolioAIEditApplyMode } from "./types";

const noteBody = (text: string): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", formatting: {}, content: [{ type: "text", text }] }],
});

/** "Term A[^1] and Term B[^2] see[^e1]." and a second paragraph. */
const buildDocx = async (): Promise<ArrayBuffer> => {
  const model = fromMarkdown("Placeholder.\n\nSecond paragraph.");
  const first = model.package.document.content[0];
  if (first?.type !== "paragraph") throw new Error("no fixture paragraph");
  first.content = [
    {
      type: "run",
      formatting: {},
      content: [
        { type: "text", text: "Term A" },
        { type: "footnoteRef", id: 30 },
        { type: "text", text: " and Term B" },
        { type: "footnoteRef", id: 10 },
        { type: "text", text: " see" },
        { type: "endnoteRef", id: 7 },
        { type: "text", text: "." },
      ],
    },
  ];
  model.package.footnotes = [30, 10].map((id, index): Footnote => ({
    type: "footnote",
    id,
    noteType: "normal",
    content: [noteBody(`Footnote ${String(index + 1)}.`)],
  }));
  const endnote: Endnote = {
    type: "endnote",
    id: 7,
    noteType: "normal",
    content: [noteBody("Endnote one.")],
  };
  model.package.endnotes = [endnote];
  const bytes = (await ensureParaIds(new Uint8Array(await createDocx(model))))
    .docx;
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
};

type Case = { name: string; find: string; replace: string; keeps: boolean };

const CASES: readonly Case[] = [
  {
    name: "both sides of a footnote reference",
    find: "A[^1] and",
    replace: "X[^1] or",
    keeps: true,
  },
  {
    name: "the words before a reference only",
    find: "Term A[^1]",
    replace: "Word A[^1]",
    keeps: true,
  },
  {
    name: "prose around two footnote references",
    find: "A[^1] and Term B[^2] see",
    replace: "Z[^1] plus Term Y[^2] near",
    keeps: true,
  },
  {
    name: "both sides of an endnote reference",
    find: "see[^e1].",
    replace: "look[^e1]!",
    keeps: true,
  },
  {
    name: "a replacement that drops the reference",
    find: "A[^1] and",
    replace: "A and",
    keeps: false,
  },
  {
    name: "a replacement that writes a reference of its own",
    find: "A[^1] and",
    replace: "A[^1] and[^3]",
    keeps: false,
  },
];

const MODES: readonly FolioAIEditApplyMode[] = ["direct", "tracked-changes"];

const open = (buffer: ArrayBuffer, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Reviewer",
    session,
  });

const save = (reviewer: FolioDocxReviewer, session: HeadlessReviewerSession) =>
  session === "canonical"
    ? canonicalReviewerSave(reviewer)
    : reviewer.toBuffer();

/** The references and the words of the saved body. */
const savedBody = async (buffer: ArrayBuffer) => {
  const zip = await ZipArchive.read(buffer);
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  return {
    references: [
      ...xml.matchAll(/<w:(footnote|endnote)Reference\b[^>]*\bw:id="(\d+)"/gu),
    ].map(([, kind, id]) => `${kind}:${id}`),
    text: [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)</gu)]
      .map(([, text]) => text)
      .join(""),
    deleted: [...xml.matchAll(/<w:delText(?: [^>]*)?>([^<]*)</gu)]
      .map(([, text]) => text)
      .join(""),
  };
};

const firstBlockText = (reviewer: FolioDocxReviewer): string =>
  reviewer.snapshot().blocks[0]?.text ?? "";

/** One engine's run of a case: the outcome, and each resolution's reads. */
const runCase = async (
  buffer: ArrayBuffer,
  session: HeadlessReviewerSession,
  { find, replace }: Case,
  mode: FolioAIEditApplyMode,
) => {
  const reviewer = await open(buffer, session);
  const blockId = reviewer.snapshot().blocks[0]?.id;
  if (blockId === undefined) throw new Error("no first block");
  const result = reviewer.applyDocumentOperations(
    {
      version: 1,
      mode,
      operations: [
        { id: "replace", type: "replaceInBlock", blockId, find, replace },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  const outcome = {
    applied: result.applied.length,
    skipped: result.skipped.map((skip) => skip.reason),
  };
  const pendingText = firstBlockText(reviewer);
  const pendingSaved = await save(reviewer, session);
  const pending = await savedBody(pendingSaved);
  const reopened = await open(pendingSaved, session);
  const reads = {
    pendingText,
    pending,
    reopenedText: firstBlockText(reopened),
  };
  if (mode === "direct") return { outcome, reads };
  const accepted = await open(buffer, session);
  accepted.applyDocumentOperations(
    {
      version: 1,
      mode,
      operations: [
        { id: "replace", type: "replaceInBlock", blockId, find, replace },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  accepted.acceptAll();
  const rejected = await open(pendingSaved, session);
  rejected.rejectAll();
  return {
    outcome,
    reads: {
      ...reads,
      accepted: await savedBody(await save(accepted, session)),
      rejected: await savedBody(await save(rejected, session)),
    },
  };
};

const SOURCE_REFERENCES = ["footnote:30", "footnote:10", "endnote:7"];

describe("a replacement over note references", () => {
  for (const testCase of CASES)
    for (const mode of MODES)
      test(`${mode}: ${testCase.name}`, async () => {
        const buffer = await buildDocx();
        const legacy = await runCase(buffer, "legacy", testCase, mode);
        const canonical = await runCase(buffer, "canonical", testCase, mode);
        expect(canonical.outcome).toEqual(legacy.outcome);
        if (testCase.keeps)
          expect(legacy.outcome).toEqual({ applied: 1, skipped: [] });
        else expect(legacy.outcome.applied).toBe(0);
        expect(canonical.reads).toEqual(legacy.reads);
        // Every reference is saved where it stands, whatever the resolution.
        expect(canonical.reads.pending.references).toEqual(SOURCE_REFERENCES);
        if ("accepted" in canonical.reads) {
          expect(canonical.reads.accepted.references).toEqual(
            SOURCE_REFERENCES,
          );
          expect(canonical.reads.rejected.references).toEqual(
            SOURCE_REFERENCES,
          );
        }
        expect(canonical.reads.reopenedText).toBe(canonical.reads.pendingText);
      }, 60_000);
});

test("the laws catch a segmentation that refuses every range holding a reference", async () => {
  // The test-only break: the shared segmentation answers no pairing for a
  // range with a reference, so a replacement there is refused, as the
  // canonical engine did before it segmented.
  const original = { ...noteReferences };
  mock.module("./note-references", () => ({
    ...original,
    segmentsAroundNoteReferences: (
      ...args: Parameters<typeof original.segmentsAroundNoteReferences>
    ) =>
      args[2].length > 0
        ? null
        : original.segmentsAroundNoteReferences(...args),
  }));
  const buffer = await buildDocx();
  const testCase = CASES[0];
  if (testCase === undefined) throw new Error("no case");
  try {
    const canonical = await runCase(buffer, "canonical", testCase, "direct");
    expect(canonical.outcome).not.toEqual({ applied: 1, skipped: [] });
  } finally {
    mock.module("./note-references", () => original);
  }
  const restored = await runCase(buffer, "canonical", testCase, "direct");
  expect(restored.outcome).toEqual({ applied: 1, skipped: [] });
}, 60_000);
