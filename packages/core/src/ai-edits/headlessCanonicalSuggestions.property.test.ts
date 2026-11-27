/**
 * A canonical headless reviewer stages, exports, restages and resolves
 * pending suggestions on its session.
 *
 * Each case draws a document and one public operation built from what a
 * reader sees, applied as a suggestion:
 *
 * - a suggestion that applies changes what a reader sees and is recorded;
 * - the saved package holds the suggestion as rejected: reopened, it reads
 *   as the document before the suggestion;
 * - the exported records restage on a reviewer reopened on the saved
 *   package, which then reads as the suggesting reviewer did;
 * - rejecting the suggestion gives the document before it back;
 * - accepting it adopts it as ordinary tracked changes: accepting every
 *   change then gives the words of the same operation applied directly, in
 *   paragraphs whose ids may differ: a paragraph the direct edit removes
 *   whole, accepting may join into, and a join keeps the second paragraph's
 *   id (docs/durable-block-ids.md).
 * - a paragraph put beside one whose break is pending removal lands where
 *   the same tracked insertion goes, in every mode: after the paragraph its
 *   pending join leaves (`insertionAfterPendingBreak`), or before the pending
 *   join that paragraph is part of (`insertionGap`). A paragraph named in a
 *   table cell has the new one beside its outermost table.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { isDeepStrictEqual } from "node:util";

import {
  pendingJoinDocumentArbitrary,
  reviewerDocumentArbitrary,
} from "../../../../test/generators/reviewerDocumentArbitraries";
import {
  type OperationDraw,
  operationFor,
  stagingDrawArbitrary,
} from "../../../../test/generators/reviewerOperationDraws";
import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { CanonicalSessionError } from "../canonical/canonicalSession";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
} from "../types/document";
import { FolioDocxReviewer } from "./headless";

type Reviewer = FolioDocxReviewer;

const open = (buffer: ArrayBuffer): Promise<Reviewer> =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Suggesting Reviewer",
    session: "canonical",
  });

const apply = (
  reviewer: Reviewer,
  operation: FolioDocumentOperation,
  mode: "direct" | "tracked-changes" | "suggested",
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: [operation],
      mode,
    },
    { undefinedReferences: "refuse" },
  );

/** What a reader sees, pending suggestions included. */
const reads = (reviewer: Reviewer) => ({
  text: reviewer.getContentAsText({ annotated: true }),
  changes: reviewer.getChanges().map(({ type, text, author, blockId }) => ({
    type,
    text,
    author,
    blockId,
  })),
  comments: reviewer
    .getComments()
    .map(({ text, anchoredText, author }) => ({ text, anchoredText, author })),
});

/** A reader's words, without the block ids each line starts with. */
const words = (text: string): string =>
  text.replace(/^\[[^\]\n]*\]/gmu, "").replace(/\s+/gu, "");

/** The block ids a reader lists, in reading order. */
const blockIds = (reviewer: Reviewer): string[] =>
  reviewer.getContent().map(({ id }) => id);

/** The texts of the blocks a reader lists, in reading order. */
const blockTexts = (reviewer: Reviewer): string[] =>
  reviewer.getContent().map(({ text }) => text);

type SuggestionOutcome =
  | "suggested"
  | "notAPackage"
  | "sessionRefused"
  | "noOperation"
  | "refusedOperation"
  | "resolutionRefused";

const suggestionLaws = async (
  buffer: ArrayBuffer,
  draw: OperationDraw,
): Promise<SuggestionOutcome> => {
  let reviewer: Reviewer;
  try {
    reviewer = await open(buffer);
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "sessionRefused";
    throw error;
  }
  const operation = operationFor(reviewer, draw, {
    replaceIn: "blocksWithWords",
  });
  if (operation === null) return "noOperation";
  const before = reads(reviewer);
  const idsBefore = blockIds(reviewer);
  const insertion =
    operation.type === "insertBeforeBlock" ||
    operation.type === "insertAfterBlock"
      ? operation
      : null;
  const besidePendingJoin =
    insertion !== null &&
    reviewer.getChanges().some(({ type }) => type === "paragraphMarkDeleted");
  const result = apply(reviewer, operation, "suggested");
  if (result.applied.length === 0) return "refusedOperation";
  const suggested = reads(reviewer);
  if (insertion !== null && besidePendingJoin) {
    // Beside a paragraph, it lands where the same tracked insertion does:
    // after the paragraph its pending join leaves, or in the gap before the
    // pending join that paragraph is part of. For one in a table cell, it is
    // beside the outermost table; before a table, before the pending join
    // that runs words into its first cell.
    const joining = new Set(
      before.changes.flatMap(({ type, blockId }) =>
        type === "paragraphMarkDeleted" && blockId !== null ? [blockId] : [],
      ),
    );
    const content = reviewer.getContent();
    const ids = content.map(({ id }) => id);
    const known = new Set(idsBefore);
    const at = ids.indexOf(insertion.blockId);
    expect(at).toBeGreaterThanOrEqual(0);
    const table = content[at]?.table?.outerTableIndex;
    const inTable = (index: number) =>
      table !== undefined && content[index]?.table?.outerTableIndex === table;
    const step = insertion.type === "insertBeforeBlock" ? -1 : 1;
    if (table === undefined) {
      // The same insertion tracked applies too: a suggestion is adopted as
      // tracked changes, so it stages only what tracking admits.
      const tracked = await open(buffer);
      expect(apply(tracked, operation, "tracked-changes").skipped).toEqual([]);
      expect(blockTexts(reviewer)).toEqual(blockTexts(tracked));
    } else {
      let edge = at;
      while (inTable(edge + step)) edge += step;
      while (
        table !== undefined &&
        step === -1 &&
        joining.has(ids[edge + step] ?? "")
      )
        edge += step;
      const beside = ids[edge + step];
      expect(beside !== undefined && !known.has(beside)).toBe(true);
    }
  }
  // A suggestion that applies changes what a reader sees and is recorded:
  // one that would change nothing is refused, not staged unseen.
  const records = reviewer.exportPendingSuggestions();
  expect({
    changed: !isDeepStrictEqual(suggested, before),
    recorded: records.length > 0,
  }).toEqual({ changed: true, recorded: true });
  const suggestionIds = [
    ...new Set(records.map(({ suggestionId }) => suggestionId)),
  ];

  // The file holds the suggestion as rejected.
  const saved = await reviewer.toBuffer();
  expect(reads(await open(saved))).toEqual(before);

  // The exported records restage on the saved package.
  const restaged = await open(saved);
  for (const loaded of restaged.loadPendingSuggestions(records))
    expect(loaded).toMatchObject({ status: "restaged" });
  expect(reads(restaged)).toEqual(suggested);

  // Rejecting the suggestion gives the document back.
  for (const id of suggestionIds)
    expect(restaged.rejectSuggestion(id)).toBe(true);
  expect(reads(restaged)).toEqual(before);

  // Accepting adopts it; accepting every change then is the direct edit.
  for (const id of suggestionIds)
    expect(reviewer.acceptSuggestion(id)).toBe(true);
  expect(reviewer.exportPendingSuggestions()).toEqual([]);
  const direct = await open(buffer);
  apply(direct, operation, "direct");
  try {
    const accepted = reviewer.acceptAll();
    const directAccepted = direct.acceptAll();
    if (accepted.skipped.length > 0 || directAccepted.skipped.length > 0)
      return "resolutionRefused";
  } catch (error) {
    if (CanonicalSessionError.is(error)) return "resolutionRefused";
    throw error;
  }
  expect(words(reviewer.getContentAsText())).toBe(
    words(direct.getContentAsText()),
  );
  return "suggested";
};

const ingested = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const prepared = await prepareCanonicalDocxInput(buffer);
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const packaged = async (document: Document): Promise<ArrayBuffer | null> => {
  try {
    return await ingested(await createDocx(structuredClone(document)));
  } catch {
    return null;
  }
};

const generatedSuggestions = (
  tally: Map<SuggestionOutcome, number>,
  documents: fc.Arbitrary<Document> = reviewerDocumentArbitrary,
) =>
  fc.asyncProperty(documents, stagingDrawArbitrary, async (document, draw) => {
    const buffer = await packaged(document);
    const outcome =
      buffer === null ? "notAPackage" : await suggestionLaws(buffer, draw);
    tally.set(outcome, (tally.get(outcome) ?? 0) + 1);
  });

/**
 * A suggestion inserted beside a paragraph whose mark is a pending insertion
 * and one whose mark is a pending deletion, in a document holding a comment
 * with no anchor (a generated case, shrunk).
 */
const besidePendingMarks: Document = {
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [],
          pPrMark: { kind: "ins", info: { id: 1, author: "A" } },
          paraId: "00000001",
        },
        {
          type: "paragraph",
          content: [],
          pPrMark: { kind: "del", info: { id: 2, author: "A" } },
          formatting: { numPr: { kind: "reference", numId: 1, ilvl: 0 } },
          paraId: "00000002",
        },
        { type: "paragraph", content: [], paraId: "00000003" },
      ],
      finalSectionProperties: { pageWidth: 12240, pageHeight: 15840 },
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
};

const commentRange: ParagraphContent[] = [
  { type: "commentRangeStart", id: 8 },
  { type: "commentRangeEnd", id: 8 },
];

/** Paragraph `n` holding `content`, under the pending changes named. */
const pending = (
  n: number,
  content: ParagraphContent[],
  mark?: "ins" | "del",
  changed = false,
): Paragraph => {
  const paragraph: Paragraph = {
    type: "paragraph",
    content,
    paraId: `0000000${n}`,
  };
  if (changed)
    paragraph.propertyChanges = [
      {
        type: "paragraphPropertyChange",
        info: { id: 10 + n, author: "A" },
        previousFormatting: { alignment: "end" },
      },
    ];
  if (mark !== undefined)
    paragraph.pPrMark = { kind: mark, info: { id: n, author: "A" } };
  return paragraph;
};

const blockBookmark: BlockContent = {
  type: "bookmarkStart",
  id: 90,
  name: "_GoBack",
};
const altChunk: BlockContent = {
  type: "preservedBlock",
  xml: "<w:altChunk/>",
};
const oneCellTable = (n: number): BlockContent => ({
  type: "table",
  columnWidths: [2400],
  rows: [
    {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [pending(n, [])] }],
    },
  ],
});

const noteReference: ParagraphContent[] = [
  { type: "run", content: [{ type: "footnoteRef", id: 2 }] },
];
const alpha: ParagraphContent[] = [
  { type: "run", content: [{ type: "text", text: "Alpha" }] },
];
/** `paragraph` as an item of list `numId`; list 2 restarts the numbering. */
const numbered = (paragraph: Paragraph, numId = 1): Paragraph => ({
  ...paragraph,
  formatting: { numPr: { kind: "reference", numId, ilvl: 0 } },
});
const groupControl = (paragraph: Paragraph): BlockContent => ({
  type: "blockSdt",
  properties: { sdtType: "group", id: 1 },
  content: [paragraph],
});
const deleting = (block: number): OperationDraw => ({
  kind: "deleteBlock",
  block,
  word: 0,
  text: "revised",
});

/**
 * Bodies where pending joins run a numbered paragraph's words into the
 * paragraphs a deletion takes, before a block their last mark cannot cross,
 * with the draw naming that deletion (the first a generated case, shrunk).
 * Directly the paragraphs go whole and the words keep their own paragraph's
 * list numbering, so accepting the suggestion numbers them alike.
 */
const DELETIONS_BESIDE_NUMBERED_JOINS: Record<
  string,
  [BlockContent[], OperationDraw]
> = {
  "deleting an empty paragraph a numbered note reference joins, before a control":
    [
      [
        numbered(pending(1, noteReference, "del")),
        pending(2, []),
        blockBookmark,
        groupControl(pending(3, [])),
        pending(4, []),
      ],
      deleting(1),
    ],
  "deleting an empty paragraph under another's added mark that a numbered note reference joins":
    [
      [
        numbered(pending(1, noteReference, "del")),
        pending(2, [], "ins"),
        blockBookmark,
        groupControl(pending(3, [])),
        pending(4, []),
      ],
      deleting(1),
    ],
  "deleting an empty paragraph before a numbered one": [
    [
      pending(1, alpha, "del"),
      pending(2, []),
      numbered(pending(3, alpha)),
      altChunk,
      pending(4, []),
    ],
    deleting(1),
  ],
  "deleting a numbered paragraph holding a note reference that words join, before a preserved block":
    [
      [
        pending(1, alpha, "del"),
        numbered(pending(2, noteReference)),
        altChunk,
        pending(3, []),
      ],
      deleting(1),
    ],
  "deleting an empty paragraph a numbered one joins, in a list that goes on after the block":
    [
      [
        numbered(pending(1, alpha)),
        numbered(pending(2, noteReference, "del")),
        pending(3, []),
        altChunk,
        numbered(pending(4, alpha)),
      ],
      deleting(2),
    ],
  "deleting an empty paragraph a numbered one joins, before a list restarting after the block":
    [
      [
        numbered(pending(1, alpha)),
        numbered(pending(2, noteReference, "del")),
        pending(3, []),
        altChunk,
        numbered(pending(4, alpha), 2),
      ],
      deleting(2),
    ],
};

const inserting = (
  kind: "insertBeforeBlock" | "insertAfterBlock",
  block: number,
): OperationDraw => ({ kind, block, word: 0, text: "revised" });

/**
 * Bodies around pending paragraph joins, with the draw naming a paragraph
 * put beside one: generated cases, shrunk, whose suggestion staged while the
 * same tracked insertion was refused. A comment's range and a property
 * change stand for any marker and property a paragraph keeps.
 */
const BESIDE_PENDING_JOINS: Record<string, [BlockContent[], OperationDraw]> = {
  "before a paragraph holding a range after a block bookmark, before a preserved block":
    [
      [
        pending(1, [], "del"),
        blockBookmark,
        pending(2, commentRange, "ins", true),
        altChunk,
      ],
      inserting("insertBeforeBlock", 1),
    ],
  "before a paragraph holding a range after a block bookmark, ending the body":
    [
      [
        pending(1, [], "del"),
        blockBookmark,
        pending(2, commentRange, "ins", true),
      ],
      inserting("insertBeforeBlock", 1),
    ],
  "after a paragraph holding a range after a block bookmark, whose added mark the new one takes over":
    [
      [
        pending(1, [], "del"),
        blockBookmark,
        pending(2, commentRange, "ins", true),
        altChunk,
        blockBookmark,
      ],
      inserting("insertAfterBlock", 1),
    ],
  "before a paragraph holding a range, whose deleted mark cannot be accepted before a block bookmark and a table":
    [
      [pending(1, commentRange, "del"), blockBookmark, oneCellTable(2)],
      inserting("insertBeforeBlock", 0),
    ],
  "before a paragraph holding a range, whose added mark cannot be rejected before a block bookmark and a table":
    [
      [
        pending(1, commentRange, "ins"),
        blockBookmark,
        oneCellTable(2),
        pending(3, [], "del"),
        pending(4, []),
      ],
      inserting("insertBeforeBlock", 0),
    ],
  "after an empty paragraph whose added mark cannot be rejected before a table":
    [
      [
        pending(1, [], "ins"),
        oneCellTable(2),
        pending(3, [], "del"),
        blockBookmark,
      ],
      inserting("insertAfterBlock", 0),
    ],
};

const bodyOf = (content: BlockContent[]): Document => ({
  package: {
    document: {
      content,
      finalSectionProperties: { pageWidth: 12240, pageHeight: 15840 },
      comments: [
        {
          id: 8,
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
      nums: [
        { numId: 1, abstractNumId: 1 },
        {
          numId: 2,
          abstractNumId: 1,
          levelOverrides: [{ ilvl: 0, startOverride: 1 }],
        },
      ],
    },
    footnotes: [
      {
        type: "footnote",
        id: 2,
        content: [{ type: "paragraph", paraId: "7FFFFFF1", content: [] }],
      },
    ],
  },
});

const GENERATED_RUNS = 30;
const PENDING_JOIN_RUNS = 100;
const COVERAGE_RUNS = 30;

describe("a canonical reviewer's pending suggestions", () => {
  // A paragraph suggested beside a pending join stages as the tracked one
  // applies; a suggested deletion a numbered pending join runs into numbers
  // its words, once accepted, as the direct deletion does.
  test.each(
    Object.entries({
      ...BESIDE_PENDING_JOINS,
      ...DELETIONS_BESIDE_NUMBERED_JOINS,
    }),
  )("a suggestion %s follows the laws", async (_shape, [content, draw]) => {
    const buffer = await packaged(bodyOf(content));
    expect(buffer).not.toBeNull();
    if (buffer === null) return;
    expect(await suggestionLaws(buffer, draw)).toBe("suggested");
  });

  test("a suggestion beside pending paragraph marks restages", async () => {
    const buffer = await packaged(besidePendingMarks);
    expect(buffer).not.toBeNull();
    if (buffer === null) return;
    expect(
      await suggestionLaws(buffer, {
        kind: "insertBeforeBlock",
        block: 1,
        word: 0,
        text: "revised",
      }),
    ).toBe("suggested");
  });

  test("a suggestion after the body paragraph before a terminal table saves and restages", async () => {
    const document = {
      package: {
        document: {
          content: [
            { type: "paragraph", paraId: "00000001", content: [] },
            {
              type: "table",
              columnWidths: [2400, 2400],
              rows: [
                {
                  type: "tableRow",
                  preservedAttributes: [
                    {
                      namespace:
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                      name: "rsidTr",
                      value: "00ABCDEF",
                    },
                  ],
                  cells: [
                    {
                      type: "tableCell",
                      content: [
                        { type: "paragraph", paraId: "00000002", content: [] },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    } satisfies Document;
    const buffer = await ingested(await createDocx(document));
    // The literal draw wraps over the three blocks the session reads (the
    // body paragraph, the cell paragraph and the paragraph closing the body
    // after the table) and targets the body paragraph.
    const draw = {
      kind: "insertAfterBlock",
      block: 3,
      word: 0,
      text: "revised",
    } as const satisfies OperationDraw;
    const reviewer = await open(buffer);
    const operation = operationFor(reviewer, draw);
    expect(operation?.type).toBe("insertAfterBlock");
    if (operation?.type !== "insertAfterBlock")
      throw new TypeError("The real draw resolves its insertion");
    expect(operation.blockId).toBe("00000001");
    const before = reviewer.toDocument();
    const applied = apply(reviewer, operation, "suggested");
    expect(applied.applied).toHaveLength(1);
    expect(applied.undoHandle).not.toBeNull();
    if (applied.undoHandle === null)
      throw new TypeError("The actual insertion has its undo handle");
    expect(reviewer.undoDocumentOperations(applied.undoHandle)).toMatchObject({
      status: "undone",
    });
    expect(reviewer.toDocument()).toStrictEqual(before);
    expect(await suggestionLaws(buffer, draw)).toBe("suggested");
  });

  test(
    "generated suggestions save, restage, reject and accept as the laws say",
    async () => {
      await assertProperty(generatedSuggestions(new Map()), {
        id: "generated suggestions save, restage, reject and accept as the laws say",
        numRuns: GENERATED_RUNS,
      });
    },
    propertyTestTimeout(120_000),
  );

  test(
    "suggestions beside pending joins save, restage, reject and accept as the laws say",
    async () => {
      await assertProperty(
        generatedSuggestions(new Map(), pendingJoinDocumentArbitrary),
        {
          id: "suggestions beside pending joins save, restage, reject and accept as the laws say",
          numRuns: PENDING_JOIN_RUNS,
        },
      );
    },
    propertyTestTimeout(120_000),
  );

  test(
    "most generated cases stage a suggestion",
    async () => {
      const tally = new Map<SuggestionOutcome, number>();
      await checkCoverageSample(generatedSuggestions(tally), COVERAGE_RUNS);
      // The whole tally is in the failure, so a shortfall names its outcomes.
      expect({
        staged: (tally.get("suggested") ?? 0) >= COVERAGE_RUNS / 3,
        tally: Object.fromEntries(tally),
      }).toMatchObject({ staged: true });
    },
    propertyTestTimeout(120_000),
  );
});
