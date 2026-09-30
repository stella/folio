/**
 * A table inserted before new paragraphs can separate the appended paragraph
 * from the paragraph that would otherwise carry its inserted break.
 * The broader tracked-resolution generator applies one operation per batch;
 * this generator composes table and paragraph insertions in the same batch.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { paragraph, run, table as buildTable } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer, type FolioEditableDocumentStoryHandle } from "./headless";

const FOOTNOTE_ID = 19;
const FOOTNOTE = { type: "footnote", noteType: "normal", id: FOOTNOTE_ID } as const;
const FOOTNOTE_STORY = { type: "footnote", noteId: FOOTNOTE_ID } as const;

setDefaultTimeout(propertyTestTimeout(120_000));

const buildSeed = async (): Promise<ArrayBuffer> => {
  const document = fromMarkdown("Anchor.");
  const [anchor] = document.package.document.content;
  if (anchor?.type !== "paragraph") panic("Missing body anchor in table-resolution fixture");
  anchor.content.unshift({
    type: "run",
    content: [{ type: "footnoteRef", id: FOOTNOTE_ID }],
  });
  document.package.footnotes = [
    {
      ...FOOTNOTE,
      content: [
        {
          type: "paragraph",
          paraId: "19000001",
          content: [{ type: "run", content: [{ type: "text", text: "Note anchor." }] }],
        },
      ],
    },
  ];
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  return new Uint8Array(docx).buffer;
};

const open = async (bytes: ArrayBuffer): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });

const content = (reviewer: FolioDocxReviewer, story?: FolioEditableDocumentStoryHandle) => {
  const blocks =
    story === undefined ? reviewer.getContent() : reviewer.snapshotStory(story)?.blocks;
  if (blocks === undefined) panic("Missing story in table-resolution fixture", { story });
  return blocks.map(({ kind, text, table }) => ({ kind, text, table: table ?? null }));
};

type OperationsForOptions = {
  blockId: string;
  paragraphCount: number;
  tableRows?: number;
  numbered?: boolean;
};

const operationsFor = ({
  blockId,
  paragraphCount,
  tableRows = 1,
  numbered = true,
}: OperationsForOptions) =>
  [
    {
      id: "table",
      type: "insertTable",
      blockId,
      rows: Array.from({ length: tableRows }, () => ["Inserted cell"]),
    },
    ...(paragraphCount === 0
      ? []
      : [
          {
            id: "paragraphs",
            type: "insertAfterBlock" as const,
            blockId,
            text: Array.from(
              { length: paragraphCount },
              (_, index) => `Inserted paragraph ${String(index)}.`,
            ).join("\n"),
            formattingScope: "allParagraphs" as const,
            ...(numbered && { numbering: { start: "new" as const, kind: "numbered" as const } }),
          },
        ]),
  ] satisfies FolioDocumentOperation[];

type ApplyOptions = {
  reviewer: FolioDocxReviewer;
  story?: FolioEditableDocumentStoryHandle;
  mode: "direct" | "tracked-changes";
  operations: FolioDocumentOperation[];
};

const apply = ({ reviewer, story, mode, operations }: ApplyOptions): void => {
  const batch = { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations };
  const result =
    story === undefined
      ? reviewer.applyDocumentOperations(batch)
      : reviewer.applyDocumentOperationsToStory({ story, batch });
  expect(result.applied.map(({ id }) => id).toSorted()).toEqual(
    operations.map(({ id }) => id).toSorted(),
  );
};

const nestedCellSeed = async (mode: "original" | "pending" | "direct"): Promise<ArrayBuffer> => {
  const document = fromMarkdown("Outside.");
  const outerTable = buildTable({ rows: [["placeholder"]] });
  const outerRow = outerTable.rows.at(0);
  const outerCell = outerRow?.cells.at(0);
  if (!outerRow || !outerCell) panic("Missing outer cell in nested-table fixture");

  const anchor = paragraph("Cell anchor.");
  const nestedTable = buildTable({ rows: [["Nested table text"]] });
  const nestedRow = nestedTable.rows.at(0);
  const nestedCell = nestedRow?.cells.at(0);
  if (!nestedRow || !nestedCell) panic("Missing nested cell in nested-table fixture");

  if (mode === "pending") {
    anchor.pPrMark = { kind: "ins", info: { id: 410, author: "Reviewer" } };
    nestedRow.structuralChange = {
      type: "tableRowInsertion",
      info: { id: 411, author: "Reviewer" },
    };
    nestedCell.content = [
      paragraph([
        {
          type: "insertion",
          info: { id: 411, author: "Reviewer" },
          content: [run("Nested table text")],
        },
      ]),
    ];
    outerCell.content = [
      anchor,
      nestedTable,
      paragraph([
        {
          type: "insertion",
          info: { id: 412, author: "Reviewer" },
          content: [run("Inserted cell paragraph.")],
        },
      ]),
    ];
  } else if (mode === "direct") {
    outerCell.content = [anchor, nestedTable, paragraph("Inserted cell paragraph.")];
  } else {
    outerCell.content = [anchor];
  }

  document.package.document.content = [outerTable, paragraph("Body end.")];
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  return new Uint8Array(docx).buffer;
};

const existingTableSeed = async (): Promise<ArrayBuffer> => {
  const document = fromMarkdown("Before table.");
  document.package.document.content.push(buildTable({ rows: [["Existing cell"]] }));
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  return new Uint8Array(docx).buffer;
};

describe("resolving inserted paragraphs after an inserted table", () => {
  for (const followingParagraphs of [0, 1, 2]) {
    test(`table plus ${String(followingParagraphs)} following paragraph(s) resolve across save/reopen`, async () => {
      const bytes = await buildSeed();
      const tracked = await open(bytes);
      const anchor = tracked.getContent().at(-1);
      if (!anchor) panic("Missing body anchor in table-resolution fixture");
      const operations = operationsFor({ blockId: anchor.id, paragraphCount: followingParagraphs });
      apply({ reviewer: tracked, mode: "tracked-changes", operations });
      const pending = content(tracked);
      if (followingParagraphs > 0) {
        expect(pending.at(-1)?.kind).toBe("listItem");
      } else {
        expect(pending.some(({ text, table }) => text === "Inserted cell" && table !== null)).toBe(
          true,
        );
      }

      const pendingReopened = await open(await tracked.toBuffer());
      expect(content(pendingReopened)).toEqual(pending);

      const rejecting = await open(await tracked.toBuffer());
      rejecting.rejectAll();
      expect(content(rejecting)).toEqual(content(await open(bytes)));
      expect(content(await open(await rejecting.toBuffer()))).toEqual(content(await open(bytes)));

      const liveRejecting = await open(bytes);
      const liveRejectAnchor = liveRejecting.getContent().at(-1);
      if (!liveRejectAnchor) panic("Missing live-reject body anchor");
      apply({
        reviewer: liveRejecting,
        mode: "tracked-changes",
        operations: operationsFor({
          blockId: liveRejectAnchor.id,
          paragraphCount: followingParagraphs,
        }),
      });
      liveRejecting.rejectAll();
      expect(content(liveRejecting)).toEqual(content(await open(bytes)));

      const direct = await open(bytes);
      const directAnchor = direct.getContent().at(-1);
      if (!directAnchor) panic("Missing direct body anchor");
      apply({
        reviewer: direct,
        mode: "direct",
        operations: operationsFor({
          blockId: directAnchor.id,
          paragraphCount: followingParagraphs,
        }),
      });

      const accepting = await open(await tracked.toBuffer());
      accepting.acceptAll();
      expect(content(accepting)).toEqual(content(direct));
      expect(content(await open(await accepting.toBuffer()))).toEqual(content(direct));

      tracked.acceptAll();
      expect(content(tracked)).toEqual(content(direct));
      expect(content(await open(await tracked.toBuffer()))).toEqual(content(direct));
    });
  }

  test("reject restores and accept matches direct when the story ends in a footnote", async () => {
    const bytes = await buildSeed();
    const tracked = await open(bytes);
    const note = tracked.snapshotStory(FOOTNOTE_STORY);
    const anchor = note?.blocks.at(-1);
    if (!anchor) panic("Missing footnote anchor in table-resolution fixture");
    const original = content(tracked, FOOTNOTE_STORY);
    const operations = operationsFor({ blockId: anchor.id, paragraphCount: 1, numbered: false });
    apply({ reviewer: tracked, story: FOOTNOTE_STORY, mode: "tracked-changes", operations });
    const pending = content(tracked, FOOTNOTE_STORY);
    expect(pending.at(-1)?.kind).toBe("paragraph");
    expect(content(await open(await tracked.toBuffer()), FOOTNOTE_STORY)).toEqual(pending);

    const rejecting = await open(await tracked.toBuffer());
    rejecting.rejectAll();
    expect(content(rejecting, FOOTNOTE_STORY)).toEqual(original);
    expect(content(await open(await rejecting.toBuffer()), FOOTNOTE_STORY)).toEqual(original);

    const direct = await open(bytes);
    const directAnchor = direct.snapshotStory(FOOTNOTE_STORY)?.blocks.at(-1);
    if (!directAnchor) panic("Missing direct footnote anchor");
    apply({
      reviewer: direct,
      story: FOOTNOTE_STORY,
      mode: "direct",
      operations: operationsFor({
        blockId: directAnchor.id,
        paragraphCount: 1,
        numbered: false,
      }),
    });

    const accepting = await open(await tracked.toBuffer());
    accepting.acceptAll();
    expect(content(accepting, FOOTNOTE_STORY)).toEqual(content(direct, FOOTNOTE_STORY));
    expect(content(await open(await accepting.toBuffer()), FOOTNOTE_STORY)).toEqual(
      content(direct, FOOTNOTE_STORY),
    );
  });

  test("a nested cell ending resolves a rotated break across an inserted table", async () => {
    const original = await open(await nestedCellSeed("original"));
    const pending = await open(await nestedCellSeed("pending"));
    const direct = await open(await nestedCellSeed("direct"));
    const pendingBuffer = await pending.toBuffer();
    expect(content(await open(pendingBuffer))).toEqual(content(pending));

    const rejecting = await open(pendingBuffer);
    rejecting.rejectAll();
    expect(content(rejecting)).toEqual(content(original));
    expect(content(await open(await rejecting.toBuffer()))).toEqual(content(original));

    const accepting = await open(pendingBuffer);
    accepting.acceptAll();
    expect(content(accepting)).toEqual(content(direct));
    expect(content(await open(await accepting.toBuffer()))).toEqual(content(direct));
  });

  test.each(["existing", "partially inserted"] as const)(
    "%s tables remain barriers for cell-anchored batches",
    async (tableState) => {
      const reviewer = await open(await existingTableSeed());
      const beforeTable = reviewer.getContent().find(({ text }) => text === "Before table.");
      if (!beforeTable) panic("Missing paragraph before existing table");
      const paragraphMarkInsertions = () =>
        reviewer
          .getChanges({ type: "paragraphMarkInserted" })
          .filter(({ blockId }) => blockId === beforeTable.id);
      expect(paragraphMarkInsertions()).toEqual([]);
      const anchor = reviewer
        .getContent()
        .find(({ text, table }) => text === "Existing cell" && table);
      if (!anchor) panic("Missing existing-table cell anchor");
      apply({
        reviewer,
        mode: "tracked-changes",
        operations: [
          ...(tableState === "partially inserted"
            ? [{ id: "row", type: "insertTableRow" as const, blockId: anchor.id }]
            : []),
          ...operationsFor({ blockId: anchor.id, paragraphCount: 1 }),
        ],
      });

      expect(paragraphMarkInsertions()).toEqual([]);
      const reopened = await open(await reviewer.toBuffer());
      expect(
        reopened
          .getChanges({ type: "paragraphMarkInserted" })
          .filter(({ blockId }) => blockId === beforeTable.id),
      ).toEqual([]);
    },
  );

  test("generated table rows and following paragraphs preserve resolution laws", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.record({
          paragraphCount: fc.integer({ min: 0, max: 3 }),
          tableRows: fc.integer({ min: 1, max: 3 }),
          numbered: fc.boolean(),
        }),
        async ({ paragraphCount, tableRows, numbered }) => {
          const bytes = await buildSeed();
          const original = await open(bytes);
          const tracked = await open(bytes);
          const anchor = tracked.getContent().at(-1);
          if (!anchor) panic("Missing generated body anchor");
          apply({
            reviewer: tracked,
            mode: "tracked-changes",
            operations: operationsFor({
              blockId: anchor.id,
              paragraphCount,
              tableRows,
              numbered,
            }),
          });

          const pendingBuffer = await tracked.toBuffer();
          const pending = content(tracked);
          expect(content(await open(pendingBuffer))).toEqual(pending);

          const rejecting = await open(pendingBuffer);
          rejecting.rejectAll();
          expect(content(rejecting)).toEqual(content(original));
          expect(content(await open(await rejecting.toBuffer()))).toEqual(content(original));

          const direct = await open(bytes);
          const directAnchor = direct.getContent().at(-1);
          if (!directAnchor) panic("Missing generated direct body anchor");
          apply({
            reviewer: direct,
            mode: "direct",
            operations: operationsFor({
              blockId: directAnchor.id,
              paragraphCount,
              tableRows,
              numbered,
            }),
          });

          const accepting = await open(pendingBuffer);
          accepting.acceptAll();
          expect(content(accepting)).toEqual(content(direct));
          expect(content(await open(await accepting.toBuffer()))).toEqual(content(direct));
        },
      ),
      { numRuns: 10 },
    );
  });
});
