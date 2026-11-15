/**
 * Undo of a committed public operation on the canonical reviewer gives back
 * the exact prior document, notes the operation removed with their
 * references included.
 *
 * The replayed document numbers a paragraph whose insertion wraps a deleted
 * footnote reference and inline math; deleting that paragraph directly
 * removes the note with it, and undo restores both.
 */

import { describe, expect, test } from "bun:test";

import {
  bodyBlockIds,
  operationFor,
} from "../../../../test/generators/reviewerOperationDraws";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const numbered = { numPr: { kind: "reference", numId: 1, ilvl: 0 } } as const;

const BODY = [
  { type: "paragraph", content: [], paraId: "00000001" },
  {
    type: "paragraph",
    content: [
      {
        type: "insertion",
        info: { id: 1, author: "A" },
        content: [
          {
            type: "deletion",
            info: { id: 5000, author: "B", date: "2026-01-02T03:04:05Z" },
            content: [
              { type: "run", content: [{ type: "footnoteRef", id: 2 }] },
            ],
          },
          { type: "mathEquation", display: "inline", ommlXml: "<m:oMath/>" },
        ],
      },
    ],
    formatting: numbered,
    paraId: "00000002",
  },
  {
    type: "paragraph",
    content: [],
    formatting: numbered,
    paraId: "00000003",
  },
] as unknown as BlockContent[];

const PAGE = { pageWidth: 12240, pageHeight: 15840 };

/** The generated document the edits property drew (seed 1462115757). */
const DOCUMENT = {
  package: {
    document: {
      content: BODY,
      sections: [{ properties: PAGE, content: BODY }],
      finalSectionProperties: PAGE,
      comments: [
        {
          id: 7,
          author: "A",
          content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
        },
      ],
    },
    numbering: {
      abstractNums: [
        {
          abstractNumId: 1,
          levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }],
        },
      ],
      nums: [{ numId: 1, abstractNumId: 1 }],
    },
    footnotes: [
      {
        type: "footnote",
        id: 2,
        content: [{ type: "paragraph", paraId: "7FFFFFF1", content: [] }],
      },
    ],
    settings: { defaultTabStop: 720 },
    properties: { title: "synthetic" },
  },
  warnings: ["kept"],
} as unknown as Document;

const packaged = async (document: Document): Promise<ArrayBuffer> => {
  const prepared = await prepareCanonicalDocxInput(
    await createDocx(structuredClone(document)),
  );
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const open = (buffer: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Canonical Reviewer",
    session: "canonical",
  });

/** Every read an undo must give back, and the document itself. */
const reads = (reviewer: FolioDocxReviewer) => ({
  content: reviewer.getContent(),
  text: reviewer.getContentAsText({ annotated: true }),
  notes: reviewer.getNotesAsText(),
  stories: reviewer.listStories(),
  changes: reviewer.getChanges(),
  comments: reviewer.getComments(),
  document: reviewer.toDocument(),
});

const applyAndUndo = async (operation: FolioDocumentOperation) => {
  const reviewer = await open(await packaged(DOCUMENT));
  const before = reads(reviewer);
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: [operation],
      mode: "direct",
    },
    { undefinedReferences: "refuse" },
  );
  return { reviewer, before, result };
};

describe("undo of a direct deletion of a paragraph holding a note reference", () => {
  test("gives back the paragraph, the note and every read", async () => {
    const probe = await open(await packaged(DOCUMENT));
    const operation = operationFor(
      probe,
      { kind: "deleteBlock", block: 1_073_741_824, word: 0, text: "revised" },
      { among: bodyBlockIds(probe.toDocument().package.document.content) },
    );
    if (operation === null) throw new Error("the draw names no block");
    const { reviewer, before, result } = await applyAndUndo(operation);
    expect(result.skipped).toEqual([]);
    if (result.undoHandle === null) throw new Error("no undo handle");
    expect(reviewer.undoDocumentOperations(result.undoHandle)).toEqual({
      status: "undone",
      undoHandle: result.undoHandle,
    });
    expect(reads(reviewer)).toEqual(before);
  }, 60_000);
});
