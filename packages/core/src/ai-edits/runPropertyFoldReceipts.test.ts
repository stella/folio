/**
 * Receipts and refusals of tracked formatting over a run's pending property
 * change. A fold over the whole run keeps the pending change's id and
 * reports it, a piece the patch cuts off takes a new one, and a batch never
 * reports an id no revision carries once it settles. A suggestion that would
 * take over another pending property change is refused by its own reason.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationResult,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";
import type { FolioAIInlineFormattingPatch } from "./types";

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const TEXT = "Formatting target";

const source = async (): Promise<ArrayBuffer> => {
  const document: Document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000001",
      content: [{ type: "run", content: [{ type: "text", text: TEXT }] }],
    },
  ];
  return await createDocx(document);
};

/** One `formatRange` over `[start, end)` of the paragraph. */
type Patch = readonly [
  start: number,
  end: number,
  formatting: FolioAIInlineFormattingPatch,
];

type BatchOptions = {
  session: HeadlessReviewerSession;
  patches: readonly Patch[];
};

const formatBatch = async ({ session, patches }: BatchOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
    author: "Reviewer",
    session,
  });
  const block = reviewer.snapshot().blocks.at(0) ?? panic("Expected a block");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: patches.map(
        ([startOffset, endOffset, formatting], index) => ({
          id: `format-${index}`,
          type: "formatRange",
          range:
            createFolioAITextRangeHandle({
              blockId: block.id,
              text: block.text,
              startOffset,
              endOffset,
            }) ?? panic("Expected a range"),
          formatting,
        }),
      ),
    },
    {
      undefinedReferences: "refuse",
      revisionStamp: { date: "2026-01-02T03:04:05.000Z", idSeed: 10 },
    },
  );
  return { reviewer, result };
};

/** The revision ids the saved package carries. */
const savedRevisionIds = async (
  reviewer: FolioDocxReviewer,
): Promise<number[]> => {
  const xml =
    (await (
      await ZipArchive.read(await reviewer.toBuffer())
    )
      .entry("word/document.xml")
      ?.text()) ?? panic("Expected word/document.xml");
  return [
    ...xml.matchAll(
      /<w:(?:ins|del|rPrChange|pPrChange)\b[^>]*\bw:id="(\d+)"/gu,
    ),
  ].map(([, id]) => Number(id));
};

const reportedIds = (result: FolioDocumentOperationResult): number[] =>
  result.applied.flatMap(({ revisionIds }) => revisionIds ?? []);

const BOLD = { bold: true } as const;
const ITALIC = { italic: true } as const;
const END = "Formatting".length;

/** Batches that format one run more than once. */
const BATCHES = {
  "bold, then italic over the same run": [
    [0, END, BOLD],
    [0, END, ITALIC],
  ],
  "bold, then italic over part of it": [
    [0, END, BOLD],
    [0, 4, ITALIC],
  ],
  "bold, then italic inside it": [
    [0, END, BOLD],
    [2, 6, ITALIC],
  ],
  "bold, then italic past it": [
    [0, END, BOLD],
    [0, TEXT.length, ITALIC],
  ],
  "italic, then bold on both halves": [
    [0, END, ITALIC],
    [0, 4, BOLD],
    [4, END, BOLD],
  ],
  "bold inside, then bold cleared around it": [
    [4, END, BOLD],
    [0, TEXT.length, { bold: null }],
  ],
  "bold inside, then italic around it": [
    [4, END, BOLD],
    [0, TEXT.length, ITALIC],
  ],
} as const satisfies Record<string, readonly Patch[]>;

describe.each(SESSIONS)("%s: a batch of run formatting", (session) => {
  test.each(Object.entries(BATCHES))(
    "%s reports only revision ids the saved package carries",
    async (_, patches) => {
      const { reviewer, result } = await formatBatch({ session, patches });
      const saved = new Set(await savedRevisionIds(reviewer));
      expect(reportedIds(result).filter((id) => !saved.has(id))).toEqual([]);
      for (const { revisionId, revisionIds } of result.applied)
        expect(revisionId).toBe(revisionIds?.at(0));
    },
  );
});

describe("canonical: ids of a folded run property change", () => {
  test("a fold over the whole run keeps the change the batch made first", async () => {
    const { reviewer, result } = await formatBatch({
      session: "canonical",
      patches: BATCHES["bold, then italic over the same run"],
    });
    expect(result.skipped).toEqual([]);
    const saved = await savedRevisionIds(reviewer);
    expect(saved).toHaveLength(1);
    // Both operations report that one change: the italic shares it.
    expect(result.applied).toEqual(
      ["format-0", "format-1"].map((id) => ({
        id,
        revisionId: saved.at(0),
        revisionIds: saved,
      })),
    );
  });

  test("a piece the patch cuts off takes a new id beside the kept one", async () => {
    const { reviewer, result } = await formatBatch({
      session: "canonical",
      patches: BATCHES["bold, then italic inside it"],
    });
    const saved = await savedRevisionIds(reviewer);
    // "Fo", "rmat" and "ting": three changes, each with its own id.
    expect(saved).toHaveLength(3);
    expect(new Set(saved).size).toBe(3);
    expect(new Set(reportedIds(result))).toEqual(new Set(saved));
  });

  test("a change a later operation settles away is reported by no receipt", async () => {
    const { reviewer, result } = await formatBatch({
      session: "canonical",
      patches: BATCHES["bold inside, then bold cleared around it"],
    });
    expect(result.skipped).toEqual([]);
    expect(await savedRevisionIds(reviewer)).toEqual([]);
    expect(result.applied.map(({ id }) => id).toSorted()).toEqual([
      "format-0",
      "format-1",
    ]);
    for (const receipt of result.applied) {
      expect(receipt).not.toHaveProperty("revisionId");
      expect(receipt).not.toHaveProperty("revisionIds");
    }
  });
});

/** "Reviewer" centred the paragraph, which stated no alignment, as change 7. */
const pendingParagraphSource = async (): Promise<ArrayBuffer> => {
  const document: Document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000001",
      formatting: { alignment: "center" },
      propertyChanges: [
        {
          type: "paragraphPropertyChange",
          info: { id: 7, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
          previousFormatting: {},
        },
      ],
      content: [{ type: "run", content: [{ type: "text", text: TEXT }] }],
    },
  ];
  return await createDocx(document);
};

/** How the reviewer answers a suggestion over another pending paragraph change. */
const PARAGRAPH_SUGGESTION = {
  canonical: {
    skipped: [
      {
        id: "align",
        reason: "pendingParagraphPropertyChange",
        message:
          "A suggestion cannot take over another pending paragraph property change.",
      },
    ],
  },
} as const satisfies Record<
  HeadlessReviewerSession,
  { skipped: readonly object[] }
>;

describe.each(SESSIONS)(
  "%s: a suggestion over another pending paragraph property change",
  (session) => {
    test("is answered by the engine's own outcome, never a missing capability", async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await pendingParagraphSource(),
        { author: "Editor", session },
      );
      const block =
        reviewer.snapshot().blocks.at(0) ?? panic("Expected a block");
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "suggested",
          operations: [
            {
              id: "align",
              type: "setBlockParagraphProperties",
              blockId: block.id,
              properties: { alignment: "right" },
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([
        ...PARAGRAPH_SUGGESTION[session].skipped,
      ]);
    });
  },
);

type ReceiptOutcome = Pick<FolioDocumentOperationResult, "applied" | "skipped">;

/**
 * A later batch's alignment and italic over the pending alignment and bold an
 * earlier batch of the same author made, given that batch's receipts. The
 * canonical session folds each edit into the pending change and reports the
 * id it shares.
 */
const LATER_BATCH = {
  canonical: (earlier: ReceiptOutcome) => ({
    applied: earlier.applied.map((receipt) => ({
      ...receipt,
      id: receipt.id.replace("-0", "-1"),
    })),
    skipped: [],
  }),
} as const satisfies Record<
  HeadlessReviewerSession,
  (earlier: ReceiptOutcome) => ReceiptOutcome
>;

type LaterBatchOptions = {
  index: number;
  alignment: "center" | "right";
  formatting: FolioAIInlineFormattingPatch;
};

describe.each(SESSIONS)(
  "%s: a later batch over the pending property changes an earlier one made",
  (session) => {
    test("reports the ids it shares", async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
        author: "Reviewer",
        session,
      });
      const block =
        reviewer.snapshot().blocks.at(0) ?? panic("Expected a block");
      const range =
        createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: 0,
          endOffset: END,
        }) ?? panic("Expected a range");
      const batch = ({ index, alignment, formatting }: LaterBatchOptions) =>
        reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "tracked-changes",
            operations: [
              {
                id: `align-${index}`,
                type: "setBlockParagraphProperties",
                blockId: block.id,
                properties: { alignment },
              },
              { id: `format-${index}`, type: "formatRange", range, formatting },
            ],
          },
          {
            undefinedReferences: "refuse",
            revisionStamp: {
              date: "2026-01-02T03:04:05.000Z",
              idSeed: 10 + index * 100,
            },
          },
        );
      const earlier = batch({
        index: 0,
        alignment: "center",
        formatting: BOLD,
      });
      expect(earlier.skipped).toEqual([]);
      const later = batch({ index: 1, alignment: "right", formatting: ITALIC });
      expect({
        applied: later.applied,
        skipped: later.skipped.map(({ id, reason }) => ({ id, reason })),
      }).toEqual(LATER_BATCH[session](earlier));
      const saved = new Set(await savedRevisionIds(reviewer));
      expect(reportedIds(later).filter((id) => !saved.has(id))).toEqual([]);
    });
  },
);
