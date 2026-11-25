/**
 * A tracked paragraph put after one whose mark is under review would take
 * that mark over: after a story's last paragraph whose mark is pending, or
 * before a table under review after a paragraph whose break is pending
 * deletion. Legacy applies it; the canonical session refuses it by a typed
 * reason, never a missing capability.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const OTHER = { author: "Other", date: "2026-01-01T00:00:00Z" };

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const pendingDeletedMark = (paraId: string, text: string): Paragraph => ({
  ...paragraph(paraId, text),
  pPrMark: { kind: "del", info: { id: 5, ...OTHER } },
});

const pendingTable: BlockContent = {
  type: "table",
  rows: [
    {
      type: "tableRow",
      structuralChange: {
        type: "tableRowInsertion",
        info: { id: 6, ...OTHER },
      },
      cells: [
        {
          type: "tableCell",
          content: [
            {
              ...paragraph("00000010", ""),
              content: [
                {
                  type: "insertion",
                  info: { id: 7, ...OTHER },
                  content: [
                    { type: "run", content: [{ type: "text", text: "cell" }] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    },
  ],
};

const SHAPES = {
  "after a story's last paragraph whose mark is pending deletion": [
    paragraph("00000001", "Opening clause."),
    pendingDeletedMark("00000002", "Closing clause."),
  ],
  "after a paragraph whose break is pending deletion, before a table under review":
    [
      paragraph("00000001", "Opening clause."),
      pendingDeletedMark("00000002", "Closing clause."),
      pendingTable,
      paragraph("00000004", "After the table."),
    ],
} as const satisfies Record<string, readonly BlockContent[]>;

const source = async (content: readonly BlockContent[]) => {
  const document: Document = createEmptyDocument();
  document.package.document.content = [...content];
  return await createDocx(document);
};

/** Each engine's outcome, by the operations it applied and skipped. */
const OUTCOME = {
  legacy: { applied: ["after"], skipped: [] },
  canonical: {
    applied: [],
    skipped: [
      {
        id: "after",
        reason: "pendingParagraphMarkDeletion",
        canonicalRefusal: undefined,
      },
    ],
  },
} as const satisfies Record<
  HeadlessReviewerSession,
  { applied: readonly string[]; skipped: readonly object[] }
>;

describe.each(SESSIONS)("%s: a tracked paragraph", (session) => {
  test.each(Object.entries(SHAPES))(
    "%s is applied, or refused by a typed reason",
    async (_, content) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await source(content),
        { author: "Reviewer", session },
      );
      const anchor =
        reviewer
          .snapshot()
          .blocks.find(({ text }) => text === "Closing clause.") ??
        panic("Expected the anchor");
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [
            {
              id: "after",
              type: "insertAfterBlock",
              blockId: anchor.id,
              text: "New clause.",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect({
        applied: result.applied.map(({ id }) => id),
        skipped: result.skipped.map(({ id, reason, canonicalRefusal }) => ({
          id,
          reason,
          canonicalRefusal,
        })),
      }).toEqual({
        applied: [...OUTCOME[session].applied],
        skipped: [...OUTCOME[session].skipped],
      });
    },
  );
});
