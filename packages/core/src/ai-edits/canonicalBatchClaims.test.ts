import { panic } from "better-result";
import { afterAll, beforeAll, expect, setSystemTime, test } from "bun:test";

import {
  paragraphsDocx,
  presentationOf,
  textRun,
} from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";
import { hashFolioAIBlockText } from "./snapshot";

beforeAll(() => setSystemTime(new Date("2026-01-01T12:00:00Z")));
afterAll(() => setSystemTime());

const source = "Alpha beta gamma delta";
const blockId = "10000000";

for (const mode of ["direct", "tracked-changes"] as const)
  for (const insertion of ["insertAfterBlock", "insertBeforeBlock"] as const)
    test(`${mode}: ${insertion} multiple insertions at one source boundary claim no existing text`, async () => {
      for (const session of ["canonical"] as const) {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await paragraphsDocx([[textRun(source)]]),
          { session },
        );
        const before = reviewer.toDocument();
        const result = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations: ["First", "Second", "Third"].map((text) => ({
              id: text,
              type: insertion,
              blockId,
              text,
            })),
          },
          { undefinedReferences: "refuse" },
        );
        expect(result.applied.map(({ id }) => id).toSorted()).toEqual([
          "First",
          "Second",
          "Third",
        ]);
        expect(result.skipped).toEqual([]);
        const undo = result.undoHandle ?? panic("An admitted batch needs undo");
        const pending = await reviewer.toBuffer();
        expect(reviewer.undoDocumentOperations(undo).status).toBe("undone");
        expect(reviewer.toDocument()).toEqual(before);
        const final = await FolioDocxReviewer.fromBuffer(pending, { session });
        if (mode === "tracked-changes")
          expect(final.acceptAll().skipped).toEqual([]);
        expect(final.getContent().map(({ text }) => text)).toEqual(
          insertion === "insertAfterBlock"
            ? [source, "First", "Second", "Third"]
            : ["First", "Second", "Third", source],
        );
      }
    });

for (const mode of ["direct", "tracked-changes"] as const)
  for (const order of ["split-first", "edit-first"] as const)
    for (const side of [
      "before",
      "after",
      "overlapping",
      "second-split",
    ] as const)
      test(`${mode}: ${order} ${side} split claims agree across engines`, async () => {
        const split = {
          id: "split",
          type: "splitBlock",
          blockId,
          offset: 10,
          separator: " ",
        } as const;
        const edit =
          side === "second-split"
            ? ({
                id: "edit",
                type: "splitBlock",
                blockId,
                offset: 5,
                separator: " ",
              } as const)
            : ({
                id: "edit",
                type: "replaceInBlock",
                blockId,
                find: {
                  before: "Alpha",
                  after: "delta",
                  overlapping: "beta gamma",
                  "second-split": "",
                }[side],
                replace: "Changed",
              } as const);
        const operations: FolioDocumentOperation[] =
          order === "split-first" ? [split, edit] : [edit, split];
        const conflicting = side === "overlapping" || side === "second-split";
        const disjointExpected =
          side === "before"
            ? ["Changed beta", "gamma delta"]
            : ["Alpha beta", "gamma Changed"];
        const editedExpected =
          side === "second-split"
            ? ["Alpha", "beta gamma delta"]
            : ["Alpha Changed delta"];
        const conflictExpected =
          order === "split-first"
            ? ["Alpha beta", "gamma delta"]
            : editedExpected;
        const expected = conflicting ? conflictExpected : disjointExpected;
        for (const session of ["canonical"] as const) {
          const bytes = await paragraphsDocx([[textRun(source)]]);
          const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
            session,
          });
          const before = reviewer.toDocument();
          const result = reviewer.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode,
              operations,
            },
            { undefinedReferences: "refuse" },
          );
          expect(result.applied.map(({ id }) => id).toSorted()).toEqual(
            conflicting
              ? [order === "split-first" ? "split" : "edit"]
              : ["edit", "split"],
          );
          expect(
            result.skipped.map(({ id, reason }) => ({ id, reason })),
          ).toEqual(
            conflicting
              ? [
                  {
                    id: order === "split-first" ? "edit" : "split",
                    reason: "overlappingOperation",
                  },
                ]
              : [],
          );
          if (mode === "tracked-changes") {
            const accepted = await FolioDocxReviewer.fromBuffer(
              await reviewer.toBuffer(),
              { session },
            );
            expect(accepted.acceptAll().skipped).toEqual([]);
            expect(accepted.getContent().map(({ text }) => text)).toEqual(
              expected,
            );
          } else
            expect(reviewer.getContent().map(({ text }) => text)).toEqual(
              expected,
            );
          const undo = result.undoHandle;
          expect(undo).toBeDefined();
          if (undo === undefined)
            panic("A committed batch must have an undo handle.");
          expect(reviewer.undoDocumentOperations(undo).status).toBe("undone");
          expect(reviewer.toDocument()).toEqual(before);
        }
      });

const range = {
  type: "textRange",
  story: "main",
  blockId,
  startOffset: 0,
  endOffset: source.length,
  selectedTextHash: hashFolioAIBlockText(source),
} as const;

for (const mode of ["direct", "tracked-changes"] as const)
  for (const reverse of [false, true])
    test(`${mode}: annotations retain their snapshot coverage before moving edits (${reverse})`, async () => {
      const presentations: ReturnType<typeof presentationOf>[] = [];
      for (const session of ["canonical"] as const) {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await paragraphsDocx([[textRun(source)]]),
          { session },
        );
        const before = reviewer.toDocument();
        const operations = [
          {
            id: "format",
            type: "formatRange",
            range,
            formatting: { bold: true },
          },
          {
            id: "comment",
            type: "commentOnRange",
            range,
            comment: { text: "Batch note" },
          },
          {
            id: "edit",
            type: "replaceInBlock",
            blockId,
            find: "Alpha",
            replace: "Longer changed opening",
          },
        ] as const satisfies readonly FolioDocumentOperation[];
        const result = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations: reverse ? operations.toReversed() : [...operations],
          },
          { undefinedReferences: "refuse" },
        );
        expect(result.skipped).toEqual([]);
        expect(result.applied.map(({ id }) => id).toSorted()).toEqual([
          "comment",
          "edit",
          "format",
        ]);
        const undo = result.undoHandle ?? panic("An admitted batch needs undo");
        const pending = await reviewer.toBuffer();
        expect(reviewer.undoDocumentOperations(undo).status).toBe("undone");
        expect(reviewer.toDocument()).toEqual(before);
        const final = await FolioDocxReviewer.fromBuffer(pending, { session });
        if (mode === "tracked-changes")
          expect(final.acceptAll().skipped).toEqual([]);
        expect(final.getContent().map(({ text }) => text)).toEqual([
          "Longer changed opening beta gamma delta",
        ]);
        presentations.push(presentationOf(final.state.doc, () => "Batch note"));
      }
      expect(presentations[1]).toEqual(presentations[0]);
    });

const tableSource = `Lead paragraph.

| Left | Right |
| --- | --- |
| A1 | B1 |
| A2 | B2 |
| A3 | B3 |

Tail paragraph.`;

for (const mode of ["direct", "tracked-changes"] as const)
  for (const reverse of [false, true])
    test(`${mode}: table row insertion, anchor removal and disjoint edits use one snapshot schedule (${reverse})`, async () => {
      const bytes = (
        await ensureParaIds(
          new Uint8Array(await createDocx(fromMarkdown(tableSource))),
        )
      ).docx.slice().buffer;
      const presentations: ReturnType<typeof presentationOf>[] = [];
      for (const session of ["canonical"] as const) {
        const reviewer = await FolioDocxReviewer.fromBuffer(bytes, { session });
        const before = reviewer.toDocument();
        const idOf = (text: string) =>
          reviewer.snapshot().blocks.find((block) => block.text === text)?.id ??
          panic("A fixture paragraph is missing");
        const operations = [
          {
            id: "row",
            type: "insertTableRow",
            blockId: idOf("A2"),
            cellTexts: ["X", "Y"],
          },
          {
            id: "lead",
            type: "insertAfterBlock",
            blockId: idOf("Lead paragraph."),
            text: "Inserted before the table.",
          },
          {
            id: "delete",
            type: "deleteBlock",
            blockId: idOf("Lead paragraph."),
          },
          {
            id: "cell",
            type: "replaceInBlock",
            blockId: idOf("A1"),
            find: "A1",
            replace: "A one",
          },
          {
            id: "tail",
            type: "splitBlock",
            blockId: idOf("Tail paragraph."),
            offset: 4,
            separator: " ",
          },
        ] as const satisfies readonly FolioDocumentOperation[];
        const result = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations: reverse ? operations.toReversed() : [...operations],
          },
          { undefinedReferences: "refuse" },
        );
        expect(result.skipped).toEqual([]);
        expect(result.applied.map(({ id }) => id).toSorted()).toEqual([
          "cell",
          "delete",
          "lead",
          "row",
          "tail",
        ]);
        const undo = result.undoHandle ?? panic("An admitted batch needs undo");
        const pending = await reviewer.toBuffer();
        expect(reviewer.undoDocumentOperations(undo).status).toBe("undone");
        expect(reviewer.toDocument()).toEqual(before);
        const final = await FolioDocxReviewer.fromBuffer(pending, { session });
        if (mode === "tracked-changes")
          expect(final.acceptAll().skipped).toEqual([]);
        expect(final.getContent().map(({ text }) => text)).toEqual([
          "Inserted before the table.",
          "Left",
          "Right",
          "A one",
          "B1",
          "A2",
          "B2",
          "X",
          "Y",
          "A3",
          "B3",
          "Tail",
          "paragraph.",
        ]);
        presentations.push(presentationOf(final.state.doc));
      }
      expect(presentations[1]).toEqual(presentations[0]);
    });

for (const mode of ["direct", "tracked-changes"] as const)
  for (const reverse of [false, true])
    test(`${mode}: column batches preserve the deletion target and account for the tracked-review boundary (${reverse})`, async () => {
      const bytes = (
        await ensureParaIds(
          new Uint8Array(await createDocx(fromMarkdown(tableSource))),
        )
      ).docx.slice().buffer;
      const presentations: ReturnType<typeof presentationOf>[] = [];
      for (const session of ["canonical"] as const) {
        const reviewer = await FolioDocxReviewer.fromBuffer(bytes, { session });
        const before = reviewer.toDocument();
        const cellBlockId =
          reviewer.snapshot().blocks.find((block) => block.text === "B1")?.id ??
          panic("A fixture cell is missing");
        const operations = [
          {
            id: "insert",
            type: "insertTableColumn",
            blockId: cellBlockId,
            position: "before",
            cellTexts: ["New", "X1", "X2", "X3"],
          },
          { id: "delete", type: "deleteTableColumn", blockId: cellBlockId },
        ] as const satisfies readonly FolioDocumentOperation[];
        const result = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations: reverse ? operations.toReversed() : [...operations],
          },
          { undefinedReferences: "refuse" },
        );
        const refusedPendingColumn =
          session === "canonical" && mode === "tracked-changes";
        expect(
          result.skipped.map(({ id, reason, canonicalRefusal }) => ({
            id,
            reason,
            canonicalRefusal,
          })),
        ).toEqual(
          refusedPendingColumn
            ? [
                {
                  id: "delete",
                  reason: "pendingTableChange",
                  canonicalRefusal: undefined,
                },
              ]
            : [],
        );
        expect(result.applied.map(({ id }) => id).toSorted()).toEqual(
          refusedPendingColumn ? ["insert"] : ["delete", "insert"],
        );
        const undo = result.undoHandle ?? panic("An admitted batch needs undo");
        const pending = await reviewer.toBuffer();
        expect(reviewer.undoDocumentOperations(undo).status).toBe("undone");
        expect(reviewer.toDocument()).toEqual(before);
        const final = await FolioDocxReviewer.fromBuffer(pending, { session });
        if (mode === "tracked-changes")
          expect(final.acceptAll().skipped).toEqual([]);
        expect(final.getContent().map(({ text }) => text)).toEqual(
          refusedPendingColumn
            ? [
                "Lead paragraph.",
                "Left",
                "New",
                "Right",
                "A1",
                "X1",
                "B1",
                "A2",
                "X2",
                "B2",
                "A3",
                "X3",
                "B3",
                "Tail paragraph.",
              ]
            : [
                "Lead paragraph.",
                "Left",
                "New",
                "A1",
                "X1",
                "A2",
                "X2",
                "A3",
                "X3",
                "Tail paragraph.",
              ],
        );
        presentations.push(presentationOf(final.state.doc));
      }
      // The tracked refusal is an existing kernel capability boundary, not
      // an equivalent result: only the implemented direct path claims parity.
      if (mode === "direct") expect(presentations[1]).toEqual(presentations[0]);
    });

for (const mode of ["direct", "tracked-changes"] as const)
  for (const reverse of [false, true])
    for (const structural of ["replace", "merge"] as const)
      test(`${mode}: ${structural} preserves paragraph-property owners (${reverse})`, async () => {
        const bytes = await paragraphsDocx([
          [textRun(source)],
          [textRun("Next paragraph.")],
          [textRun("Final paragraph.")],
        ]);
        const presentations: ReturnType<typeof presentationOf>[] = [];
        for (const session of ["canonical"] as const) {
          const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
            session,
          });
          const before = reviewer.toDocument();
          const edit =
            structural === "replace"
              ? ({
                  id: "structure",
                  type: "replaceInBlock",
                  blockId,
                  find: "Alpha",
                  replace: "Changed",
                } as const)
              : ({
                  id: "structure",
                  type: "mergeBlockWithNext",
                  blockId,
                  separator: " ",
                } as const);
          const properties = {
            id: "properties",
            type: "setBlockParagraphProperties",
            blockId,
            properties: { alignment: "center" },
          } as const;
          const operations = reverse ? [properties, edit] : [edit, properties];
          const result = reviewer.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode,
              operations,
            },
            { undefinedReferences: "refuse" },
          );
          const pendingMerge =
            mode === "tracked-changes" && structural === "merge";
          expect(
            result.skipped.map(({ id, reason }) => ({ id, reason })),
            session,
          ).toEqual(
            pendingMerge
              ? [{ id: "properties", reason: "pendingParagraphMarkDeletion" }]
              : [],
          );
          expect(
            result.applied.map(({ id }) => id).toSorted(),
            session,
          ).toEqual(pendingMerge ? ["structure"] : ["properties", "structure"]);
          const undo =
            result.undoHandle ?? panic("An admitted batch needs undo");
          const pending = await reviewer.toBuffer();
          expect(reviewer.undoDocumentOperations(undo).status).toBe("undone");
          expect(reviewer.toDocument()).toEqual(before);
          const final = await FolioDocxReviewer.fromBuffer(pending, {
            session,
          });
          if (mode === "tracked-changes")
            expect(final.acceptAll().skipped).toEqual([]);
          const presentation = presentationOf(final.state.doc);
          expect(presentation.map(({ text }) => text)).toEqual(
            structural === "replace"
              ? [
                  "Changed beta gamma delta",
                  "Next paragraph.",
                  "Final paragraph.",
                ]
              : ["Alpha beta gamma delta Next paragraph.", "Final paragraph."],
          );
          expect(presentation.at(0)?.alignment).toBe(
            pendingMerge ? null : "center",
          );
          presentations.push(presentation);
        }
        expect(presentations[1]).toEqual(presentations[0]);
      });

test("canonical column review refusal preserves the already-pending table atomically", async () => {
  const bytes = (
    await ensureParaIds(
      new Uint8Array(await createDocx(fromMarkdown(tableSource))),
    )
  ).docx.slice().buffer;
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
    session: "canonical",
  });
  const original = reviewer.toDocument();
  const cellBlockId =
    reviewer.snapshot().blocks.find((block) => block.text === "B1")?.id ??
    panic("A fixture cell is missing");
  const inserted = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "insert",
          type: "insertTableColumn",
          blockId: cellBlockId,
          position: "before",
          cellTexts: ["New", "X1", "X2", "X3"],
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(inserted.skipped).toEqual([]);
  expect(inserted.applied.map(({ id }) => id)).toEqual(["insert"]);
  const undo = inserted.undoHandle ?? panic("The column insertion needs undo");
  const before = reviewer.toDocument();
  const changes = reviewer.getChanges();
  const saved = await reviewer.toBuffer();
  const deleted = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        { id: "delete", type: "deleteTableColumn", blockId: cellBlockId },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(deleted.applied).toEqual([]);
  expect(
    deleted.skipped.map(({ id, reason, canonicalRefusal }) => ({
      id,
      reason,
      canonicalRefusal,
    })),
  ).toEqual([
    {
      id: "delete",
      reason: "pendingTableChange",
      canonicalRefusal: undefined,
    },
  ]);
  expect(deleted.undoHandle).toBeNull();
  expect(reviewer.toDocument()).toEqual(before);
  expect(reviewer.getChanges()).toEqual(changes);
  expect(new Uint8Array(await reviewer.toBuffer())).toEqual(
    new Uint8Array(saved),
  );
  expect(reviewer.undoDocumentOperations(undo).status).toBe("undone");
  expect(reviewer.toDocument()).toEqual(original);
});

for (const columns of [1, 2])
  for (const reverse of [false, true])
    test(`direct column mapping counts only successful insertions (${columns} source columns, reverse=${reverse})`, async () => {
      const markdown =
        columns === 1
          ? "Lead paragraph.\n\n| Right |\n| --- |\n| B1 |\n| B2 |\n| B3 |\n\nTail paragraph."
          : tableSource;
      const bytes = (
        await ensureParaIds(
          new Uint8Array(await createDocx(fromMarkdown(markdown))),
        )
      ).docx.slice().buffer;
      const presentations: ReturnType<typeof presentationOf>[] = [];
      for (const session of ["canonical"] as const) {
        const reviewer = await FolioDocxReviewer.fromBuffer(bytes, { session });
        const before = reviewer.toDocument();
        const cellBlockId =
          reviewer.snapshot().blocks.find((block) => block.text === "B1")?.id ??
          panic("A fixture cell is missing");
        const operations = [
          {
            id: "first",
            type: "insertTableColumn",
            blockId: cellBlockId,
            position: "before",
            cellTexts: ["First", "F1", "F2", "F3"],
          },
          {
            id: "overflow",
            type: "insertTableColumn",
            blockId: cellBlockId,
            position: "before",
            cellTexts: ["Overflow", "O1", "O2", "O3", "O4"],
          },
          {
            id: "second",
            type: "insertTableColumn",
            blockId: cellBlockId,
            position: "before",
            cellTexts: ["Second", "S1", "S2", "S3"],
          },
          { id: "delete", type: "deleteTableColumn", blockId: cellBlockId },
        ] as const satisfies readonly FolioDocumentOperation[];
        const result = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "direct",
            operations: reverse ? operations.toReversed() : [...operations],
          },
          { undefinedReferences: "refuse" },
        );
        expect(result.applied.map(({ id }) => id).toSorted()).toEqual([
          "delete",
          "first",
          "second",
        ]);
        expect(
          result.skipped.map(({ id, reason }) => ({ id, reason })),
        ).toEqual([{ id: "overflow", reason: "payloadDoesNotFit" }]);
        const undo = result.undoHandle ?? panic("An admitted batch needs undo");
        const pending = await reviewer.toBuffer();
        expect(reviewer.undoDocumentOperations(undo).status).toBe("undone");
        expect(reviewer.toDocument()).toEqual(before);
        const final = await FolioDocxReviewer.fromBuffer(pending, { session });
        const inserted = reverse
          ? ["Second", "First", "S1", "F1", "S2", "F2", "S3", "F3"]
          : ["First", "Second", "F1", "S1", "F2", "S2", "F3", "S3"];
        const expected =
          columns === 1
            ? ["Lead paragraph.", ...inserted, "Tail paragraph."]
            : [
                "Lead paragraph.",
                "Left",
                ...inserted.slice(0, 2),
                "A1",
                ...inserted.slice(2, 4),
                "A2",
                ...inserted.slice(4, 6),
                "A3",
                ...inserted.slice(6),
                "Tail paragraph.",
              ];
        expect(final.getContent().map(({ text }) => text)).toEqual(expected);
        presentations.push(presentationOf(final.state.doc));
      }
      expect(presentations[1]).toEqual(presentations[0]);
    });

for (const mode of ["direct", "tracked-changes"] as const)
  for (const reverse of [false, true])
    test(`${mode}: an insertion after a block and its merge claim one gap (reverse=${reverse})`, async () => {
      const outcomes: unknown[] = [];
      for (const session of ["canonical"] as const) {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await paragraphsDocx([
            [textRun("Alpha.")],
            [textRun("Bravo.")],
            [textRun("Charlie.")],
          ]),
          { session },
        );
        const first =
          reviewer.snapshot().blocks[0]?.id ??
          panic("The fixture needs its first paragraph");
        const operations = [
          {
            id: "insert",
            type: "insertAfterBlock",
            blockId: first,
            text: "New.",
          },
          { id: "merge", type: "mergeBlockWithNext", blockId: first },
        ] as const satisfies readonly FolioDocumentOperation[];
        const result = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations: reverse ? operations.toReversed() : [...operations],
          },
          { undefinedReferences: "refuse" },
        );
        const final = await FolioDocxReviewer.fromBuffer(
          await reviewer.toBuffer(),
          { session },
        );
        if (mode === "tracked-changes")
          expect(final.acceptAll().skipped).toEqual([]);
        outcomes.push({
          skipped: result.skipped.map(({ id, reason }) => ({ id, reason })),
          texts: final.getContent().map(({ text }) => text),
        });
      }
      expect(outcomes[0]).toEqual({
        skipped: [
          {
            id: reverse ? "insert" : "merge",
            reason: "overlappingOperation",
          },
        ],
        texts: reverse
          ? ["Alpha.Bravo.", "Charlie."]
          : ["Alpha.", "New.", "Bravo.", "Charlie."],
      });
      expect(outcomes[1]).toEqual(outcomes[0]);
    });

for (const mode of ["direct", "tracked-changes"] as const)
  for (const reverse of [false, true])
    test(`${mode}: a column inserted after a one-column table's only column survives its deletion (reverse=${reverse})`, async () => {
      const bytes = (
        await ensureParaIds(
          new Uint8Array(
            await createDocx(
              fromMarkdown(
                "Lead paragraph.\n\n| Only |\n| --- |\n| B1 |\n\nTail paragraph.",
              ),
            ),
          ),
        )
      ).docx.slice().buffer;
      const outcomes: unknown[] = [];
      for (const session of ["canonical"] as const) {
        const reviewer = await FolioDocxReviewer.fromBuffer(bytes, { session });
        const cellBlockId =
          reviewer.snapshot().blocks.find((block) => block.text === "B1")?.id ??
          panic("A fixture cell is missing");
        const operations = [
          {
            id: "insert",
            type: "insertTableColumn",
            blockId: cellBlockId,
            position: "after",
            cellTexts: ["Added", "A1"],
          },
          { id: "delete", type: "deleteTableColumn", blockId: cellBlockId },
        ] as const satisfies readonly FolioDocumentOperation[];
        const result = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations: reverse ? operations.toReversed() : [...operations],
          },
          { undefinedReferences: "refuse" },
        );
        const final = await FolioDocxReviewer.fromBuffer(
          await reviewer.toBuffer(),
          { session },
        );
        if (mode === "tracked-changes")
          expect(final.acceptAll().skipped).toEqual([]);
        outcomes.push({
          applied: result.applied.map(({ id }) => id).toSorted(),
          skipped: result.skipped.map(({ id, reason }) => ({ id, reason })),
          texts: final.getContent().map(({ text }) => text),
        });
      }
      const kept = {
        applied: ["delete", "insert"],
        skipped: [],
        texts: ["Lead paragraph.", "Added", "A1", "Tail paragraph."],
      };
      // Both engines remove only the source column. Tracked, the canonical
      // session asks for the table's pending review to be resolved first and
      // leaves the table, the inserted column included, standing.
      expect(outcomes).toEqual([
        kept,
        mode === "direct"
          ? kept
          : {
              applied: ["insert"],
              skipped: [{ id: "delete", reason: "pendingTableChange" }],
              texts: [
                "Lead paragraph.",
                "Only",
                "Added",
                "B1",
                "A1",
                "Tail paragraph.",
              ],
            },
      ]);
    });
