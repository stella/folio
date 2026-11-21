/**
 * A tracked operation that would record over a paragraph mark under review
 * is refused by what the mark has pending: `pendingParagraphMarkInsertion`
 * for an inserted mark, `pendingParagraphMarkDeletion` for a deleted one, so
 * the two are never reported alike. Legacy applies these shapes.
 */
import { describe, expect, test } from "bun:test";

import { PARAGRAPH_MARK_CHANGE_KINDS } from "@stll/docx-core/model";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphMarkChangeKind,
  TableRow,
} from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIEditOperation, FolioAIEditSkipReason } from "./types";

const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const OTHER = { author: "Other", date: "2026-01-01T00:00:00Z" };

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const pendingMark = (
  kind: ParagraphMarkChangeKind,
  paraId: string,
  text: string,
): Paragraph => ({
  ...paragraph(paraId, text),
  pPrMark: { kind, info: { id: 5, ...OTHER } },
});

const row = (cells: readonly Paragraph[]): TableRow => ({
  type: "tableRow",
  cells: cells.map((cell) => ({ type: "tableCell", content: [cell] })),
});

/**
 * The canonical reason for each pending mark kind: a break inserted or moved
 * here is pending insertion, one deleted or moved away pending deletion.
 */
const CANONICAL_REASON = {
  ins: "pendingParagraphMarkInsertion",
  moveTo: "pendingParagraphMarkInsertion",
  del: "pendingParagraphMarkDeletion",
  moveFrom: "pendingParagraphMarkDeletion",
} as const satisfies Record<ParagraphMarkChangeKind, FolioAIEditSkipReason>;

type MarkKind = ParagraphMarkChangeKind;

type Shape = {
  content: (kind: MarkKind) => BlockContent[];
  /** The pending mark kinds the shape refuses by what is pending. */
  kinds: readonly MarkKind[];
  operation: FolioAIEditOperation;
  message: string;
};

const SHAPES = {
  "a paragraph after a story's last paragraph": {
    content: (kind) => [
      paragraph("00000001", "Opening."),
      pendingMark(kind, "00000002", "Closing."),
    ],
    kinds: PARAGRAPH_MARK_CHANGE_KINDS,
    operation: {
      id: "op",
      type: "insertAfterBlock",
      blockId: "00000002",
      text: "New.",
    },
    message:
      "A tracked insertion after a paragraph no paragraph follows needs the anchor's mark free of a pending change.",
  },
  "a table after a story's last paragraph": {
    content: (kind) => [
      paragraph("00000001", "Opening."),
      pendingMark(kind, "00000002", "Closing."),
    ],
    kinds: PARAGRAPH_MARK_CHANGE_KINDS,
    operation: {
      id: "op",
      type: "insertTable",
      blockId: "00000002",
      position: "after",
      rows: [["Cell"]],
    },
    message:
      "A tracked insertion after a paragraph no paragraph follows needs the anchor's mark free of a pending change.",
  },
  "a row deletion over a cell whose final paragraph's mark is pending": {
    content: (kind) => [
      paragraph("00000001", "Before."),
      {
        type: "table",
        rows: [
          row([paragraph("00000010", "A1"), paragraph("00000011", "B1")]),
          row([
            pendingMark(kind, "00000012", "A2"),
            paragraph("00000013", "B2"),
          ]),
        ],
      },
      paragraph("00000002", "After."),
    ],
    // A moving mark in a cell is a shape tracked rows do not record yet.
    kinds: ["ins", "del"],
    operation: { id: "op", type: "deleteTableRow", blockId: "00000012" },
    message: "A cell's final paragraph already carries a mark revision.",
  },
} satisfies Record<string, Shape>;

const source = async (content: BlockContent[]) => {
  const document: Document = createEmptyDocument();
  document.package.document.content = content;
  return await createDocx(document);
};

const outcome = (session: HeadlessReviewerSession, kind: MarkKind) => {
  switch (session) {
    case "legacy":
      return { applied: ["op"], skipped: [] };
    case "canonical":
      return {
        applied: [],
        skipped: [
          {
            id: "op",
            reason: CANONICAL_REASON[kind],
            canonicalRefusal: undefined,
          },
        ],
      };
    default: {
      const unreachable: never = session;
      return unreachable;
    }
  }
};

const CASES = Object.entries(SHAPES).flatMap(([name, shape]) =>
  shape.kinds.map((kind) => ({ name, kind, ...shape })),
);

describe.each(SESSIONS)(
  "%s: a tracked operation over a pending mark",
  (session) => {
    test.each(CASES)(
      "$name ($kind) is applied, or refused by what the mark has pending",
      async ({ kind, content, operation, message }) => {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await source(content(kind)),
          { author: "Reviewer", session },
        );
        const result = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "tracked-changes",
            operations: [operation],
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
        }).toEqual(outcome(session, kind));
        expect(result.skipped.map(({ message: said }) => said)).toEqual(
          session === "canonical" ? [message] : [],
        );
      },
    );
  },
);
