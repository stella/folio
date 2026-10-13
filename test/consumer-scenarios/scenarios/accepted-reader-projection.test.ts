import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  docxToMarkdown,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  parseFolioDocumentOperationBatch,
  type FolioDocumentOperation,
  type FolioDocumentOperationResult,
} from "@stll/folio-core/server";

import {
  notesDocument,
  openReviewer,
  packDocument,
  tableDocument,
  toArrayBuffer,
} from "../support/documents.ts";
import { parseFlowFile } from "../support/flow-file.ts";
import { assertReadersAgree } from "../support/invariants.ts";
import { MARKDOWN_READ_OPTIONS, readAll } from "../support/readers.ts";
import { capture } from "../support/oracle.ts";

/** Markup readers retain structural revisions; clean Markdown must match actual acceptance. */
const assertAcceptedReaders = async (bytes: Uint8Array, label: string) => {
  const pending = await readAll(bytes);
  const reviewer = await openReviewer(bytes);
  reviewer.acceptAll();
  const acceptedBytes = new Uint8Array(await reviewer.toBuffer());
  const accepted = await readAll(acceptedBytes);
  // Compare the raw document too: BlockView deliberately omits note trailers.
  assert.equal(
    await docxToMarkdown(toArrayBuffer(bytes), MARKDOWN_READ_OPTIONS),
    await docxToMarkdown(toArrayBuffer(acceptedBytes), MARKDOWN_READ_OPTIONS),
    `${label}: full clean Markdown including referenced note trailers`,
  );
  assert.deepEqual(
    pending.markdown,
    accepted.markdown,
    `${label}: clean Markdown before/after acceptance`,
  );
  assert.deepEqual(
    pending.getContentAsMarkdown,
    accepted.getContentAsMarkdown,
    `${label}: accepted block view before/after save`,
  );
  await assertReadersAgree(bytes, `${label}: pending package`);
  await assertReadersAgree(acceptedBytes, `${label}: accepted package`);
};

type InsertionBinding =
  | { type: "paragraph"; operationId: string }
  | { type: "tableCell"; operationId: string; row: number; cell: number };

// Unstamped insertions allocate new paragraph IDs. Keep the recorded payloads
// exact and bind their new IDs to the operation-produced owner, never a fallback.
const INSERTION_BINDINGS = {
  "17340089": { type: "paragraph", operationId: "op-insertBeforeBlock" },
  "425BB368": { type: "paragraph", operationId: "op-insertAfterBlock" },
  "7956AFCA": { type: "tableCell", operationId: "op-insertTable", row: 0, cell: 0 },
  "5B7407EF": { type: "tableCell", operationId: "op-insertTable", row: 1, cell: 1 },
} as const satisfies Record<string, InsertionBinding>;
type RecordedInsertionId = keyof typeof INSERTION_BINDINGS;
const isRecordedInsertionId = (id: string): id is RecordedInsertionId =>
  Object.hasOwn(INSERTION_BINDINGS, id);

type RebindOperationOptions = {
  operation: FolioDocumentOperation;
  inserted: ReadonlyMap<RecordedInsertionId, string>;
  fixtureIds: ReadonlySet<string>;
};
const rebindOperation = ({
  operation,
  inserted,
  fixtureIds,
}: RebindOperationOptions): FolioDocumentOperation => {
  const blockId = (recorded: string): string => {
    if (!isRecordedInsertionId(recorded)) {
      assert.ok(fixtureIds.has(recorded), `Unbound recorded paragraph ID ${recorded}`);
      return recorded;
    }
    const actual = inserted.get(recorded);
    assert.ok(actual, `Recorded insertion ${recorded} has not been created`);
    return actual;
  };
  switch (operation.type) {
    case "replaceRange":
    case "commentOnRange":
    case "formatRange":
      return {
        ...operation,
        range: { ...operation.range, blockId: blockId(operation.range.blockId) },
      };
    case "mergeTableCells":
      if (operation.endBlockId !== undefined) {
        return {
          ...operation,
          blockId: blockId(operation.blockId),
          endBlockId: blockId(operation.endBlockId),
        };
      }
      return { ...operation, blockId: blockId(operation.blockId) };
    case "replaceInBlock":
    case "insertAfterBlock":
    case "insertBeforeBlock":
    case "replaceBlock":
    case "deleteBlock":
    case "splitBlock":
    case "mergeBlockWithNext":
    case "setBlockParagraphProperties":
    case "insertTable":
    case "deleteTable":
    case "commentOnBlock":
    case "insertSignatureTable":
    case "insertTableRow":
    case "deleteTableRow":
    case "insertTableColumn":
    case "deleteTableColumn":
    case "splitTableCell":
      return { ...operation, blockId: blockId(operation.blockId) };
    default: {
      const unhandled: never = operation;
      throw new Error(`Unmapped reader flow operation: ${JSON.stringify(unhandled)}`);
    }
  }
};

const NOOP_PROPERTIES_STEP = 10;
const expectedNoop = [
  { id: "op-setBlockParagraphProperties", reason: "noopOperation" },
] as const satisfies readonly Pick<
  FolioDocumentOperationResult["skipped"][number],
  "id" | "reason"
>[];

test("recorded operation matrix seed 201 keeps clean readers consistent through table deletion", async () => {
  const input: unknown = JSON.parse(
    readFileSync(
      new URL("../support/reader-flows/table-deletion-seed201.json", import.meta.url),
      "utf8",
    ),
  );
  const flow = parseFlowFile(input);
  assert.equal(flow.fixture, "tables");
  const reviewer = await openReviewer(await tableDocument());
  const fixtureIds = new Set(reviewer.getContent().map(({ id }) => id));
  const inserted = new Map<RecordedInsertionId, string>();
  for (const [index, step] of flow.steps.entries()) {
    assert.equal(step.action, "core batch");
    assert.ok(step.operations);
    const batch = parseFolioDocumentOperationBatch({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: flow.mode,
      atomic: false,
      operations: step.operations,
    });
    const rebound = {
      ...batch,
      operations: batch.operations.map((operation) =>
        rebindOperation({ operation, inserted, fixtureIds }),
      ),
    };
    const beforeIds = new Set(reviewer.getContent().map(({ id }) => id));
    // Preserve the original matrix's pre-operation accepted/rejected copies;
    // revision allocation influences the inserted identities later steps name.
    await capture(reviewer, "tracked-changes");
    const result = reviewer.applyDocumentOperations(rebound, { undefinedReferences: "refuse" });
    assert.deepEqual(
      result.applied.map(({ id }) => id),
      index === NOOP_PROPERTIES_STEP ? [] : batch.operations.map(({ id }) => id),
      `seed 201 step ${index}: applied operation ids; ${JSON.stringify(result.issues)}`,
    );
    assert.deepEqual(
      result.skipped.map(({ id, reason }) => ({ id, reason })),
      index === NOOP_PROPERTIES_STEP ? expectedNoop : [],
      `seed 201 step ${index}: recorded refusals`,
    );
    for (const operation of rebound.operations) {
      const created = reviewer.getContent().filter(({ id }) => !beforeIds.has(id));
      for (const recordedId of Object.keys(INSERTION_BINDINGS)) {
        assert.ok(isRecordedInsertionId(recordedId));
        const binding = INSERTION_BINDINGS[recordedId];
        if (binding.operationId !== operation.id) continue;
        switch (binding.type) {
          case "paragraph": {
            assert.ok(
              operation.type === "insertAfterBlock" || operation.type === "insertBeforeBlock",
            );
            assert.equal(created.length, 1);
            const paragraph = created.at(0);
            assert.ok(paragraph);
            assert.equal(paragraph.table, undefined);
            assert.equal(paragraph.text, operation.text);
            inserted.set(recordedId, paragraph.id);
            break;
          }
          case "tableCell": {
            assert.ok(operation.type === "insertTable");
            assert.equal(
              created.length,
              operation.rows.reduce((total, row) => total + row.length, 0),
            );
            assert.equal(new Set(created.map(({ table }) => table?.tableIndex)).size, 1);
            const cells = created.filter(
              ({ table }) => table?.rowIndex === binding.row && table.cellIndex === binding.cell,
            );
            assert.equal(cells.length, 1);
            const cell = cells.at(0);
            assert.ok(cell);
            assert.equal(cell.text, operation.rows.at(binding.row)?.at(binding.cell));
            inserted.set(recordedId, cell.id);
            break;
          }
          default: {
            const unhandled: never = binding;
            throw new Error(`Unknown insertion binding ${JSON.stringify(unhandled)}`);
          }
        }
      }
    }
  }
  await assertAcceptedReaders(new Uint8Array(await reviewer.toBuffer()), "recorded seed 201");
});

test("clean readers equal accept-all over paragraph deletion and join positions and style directions", async () => {
  for (let count = 2; count <= 4; count++) {
    for (const headingPosition of [0, count - 1, -1]) {
      const paragraphs = Array.from(
        { length: count },
        (_, index) => `${index === headingPosition ? "# " : ""}Clause ${index}.`,
      );
      const fixture = await packDocument(fromMarkdown(paragraphs.join("\n\n")));
      for (const type of ["deleteBlock", "mergeBlockWithNext"] as const) {
        for (let position = 0; position < count; position++) {
          if (type === "mergeBlockWithNext" && position === count - 1) continue;
          const reviewer = await openReviewer(fixture);
          const block = reviewer.getContent().at(position);
          assert.ok(block);
          const result = reviewer.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode: "tracked-changes",
              operations: [
                type === "deleteBlock"
                  ? { id: "change", type, blockId: block.id }
                  : { id: "change", type, blockId: block.id, separator: " " },
              ],
            },
            { undefinedReferences: "refuse" },
          );
          assert.equal(result.applied.length, 1, JSON.stringify(result.issues));
          await assertAcceptedReaders(
            new Uint8Array(await reviewer.toBuffer()),
            `${type}, count=${count}, position=${position}, heading=${headingPosition}`,
          );
        }
      }
    }
  }
});

test("clean readers resolve table, row, and column deletions in every source position", async () => {
  for (const width of [2, 3]) {
    const tableMarkdown = [
      `| ${Array.from({ length: width }, (_, column) => `Header ${column}`).join(" | ")} |`,
      `| ${Array.from({ length: width }, () => "---").join(" | ")} |`,
      ...Array.from(
        { length: 3 },
        (_, row) =>
          `| ${Array.from({ length: width }, (_cell, column) => `Cell ${row}:${column}`).join(" | ")} |`,
      ),
    ].join("\n");
    const fixture = await packDocument(
      fromMarkdown(`Before table.\n\n${tableMarkdown}\n\nAfter table.`),
    );
    for (const type of ["deleteTable", "deleteTableRow", "deleteTableColumn"] as const) {
      const probe = await openReviewer(fixture);
      const anchors = probe.getContent().filter(({ table }) => {
        if (!table) return false;
        switch (type) {
          case "deleteTable":
            return table.rowIndex === 0 && table.cellIndex === 0;
          case "deleteTableRow":
            return table.cellIndex === 0;
          case "deleteTableColumn":
            return table.rowIndex === 0;
          default: {
            const unhandled: never = type;
            throw new Error(`Unhandled deletion ${unhandled}`);
          }
        }
      });
      const expectedAnchorCounts = {
        deleteTable: 1,
        deleteTableRow: 4,
        deleteTableColumn: width,
      } as const satisfies Record<typeof type, number>;
      assert.equal(anchors.length, expectedAnchorCounts[type]);
      for (const anchor of anchors) {
        const reviewer = await openReviewer(fixture);
        const result = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "tracked-changes",
            operations: [{ id: "delete", type, blockId: anchor.id }],
          },
          { undefinedReferences: "refuse" },
        );
        assert.equal(result.applied.length, 1, JSON.stringify(result.issues));
        await assertAcceptedReaders(
          new Uint8Array(await reviewer.toBuffer()),
          `${type}, width=${width}, row=${anchor.table?.rowIndex}, cell=${anchor.table?.cellIndex}`,
        );
      }
    }
    for (const order of ["paragraph-first", "table-first"] as const) {
      const reviewer = await openReviewer(fixture);
      const before = reviewer.getContent().at(0);
      const cell = reviewer.getContent().find(({ table }) => table !== undefined);
      assert.ok(before && cell);
      const paragraphDeletion = {
        id: "paragraph",
        type: "deleteBlock",
        blockId: before.id,
      } as const satisfies FolioDocumentOperation;
      const tableDeletion = {
        id: "table",
        type: "deleteTable",
        blockId: cell.id,
      } as const satisfies FolioDocumentOperation;
      const operations =
        order === "paragraph-first"
          ? [paragraphDeletion, tableDeletion]
          : [tableDeletion, paragraphDeletion];
      for (const operation of operations) {
        const result = reviewer.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "tracked-changes",
            operations: [operation],
          },
          { undefinedReferences: "refuse" },
        );
        assert.equal(result.applied.length, 1, JSON.stringify(result.issues));
      }
      await assertAcceptedReaders(
        new Uint8Array(await reviewer.toBuffer()),
        `paragraph/table deletion, width=${width}, order=${order}`,
      );
    }
  }
});

test("clean Markdown keeps referenced footnote and endnote trailers through body joins and deletions", async () => {
  const fixture = await notesDocument();
  for (const prefix of ["Goods are defined", "Warranty terms apply"]) {
    for (const type of ["deleteBlock", "mergeBlockWithNext"] as const) {
      const reviewer = await openReviewer(fixture);
      const blocks = reviewer.getContent();
      const position = blocks.findIndex(({ text }) => text.startsWith(prefix));
      assert.ok(position > 0);
      const block = blocks.at(type === "mergeBlockWithNext" ? position - 1 : position);
      assert.ok(block);
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [
            type === "deleteBlock"
              ? { id: "change", type, blockId: block.id }
              : { id: "change", type, blockId: block.id, separator: " " },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      assert.equal(result.applied.length, 1, JSON.stringify(result.issues));
      await assertAcceptedReaders(
        new Uint8Array(await reviewer.toBuffer()),
        `${type}: reference ${prefix}`,
      );
    }
  }
});
test("referenced footnote and endnote trailers resolve paragraph joins and table deletions", async () => {
  const document = (await openReviewer(await notesDocument())).toDocument();
  const noteContent = fromMarkdown(
    "Note lead.\n\nNote join.\n\n| Term | Value |\n| --- | --- |\n| Goods | Paid |\n\nNote ending.",
  ).package.document.content;
  for (const note of [
    ...(document.package.footnotes ?? []),
    ...(document.package.endnotes ?? []),
  ]) {
    if (note.id > 0) note.content = structuredClone(noteContent);
  }
  const fixture = await packDocument(document);
  for (const story of [
    { type: "footnote", noteId: 1 },
    { type: "endnote", noteId: 1 },
  ] as const) {
    for (const type of ["deleteBlock", "mergeBlockWithNext", "deleteTable"] as const) {
      const reviewer = await openReviewer(fixture);
      const snapshot = reviewer.snapshotStory(story);
      assert.ok(snapshot);
      const block =
        type === "deleteTable"
          ? snapshot.blocks.find(({ table }) => table !== undefined)
          : snapshot.blocks.at(0);
      assert.ok(block);
      const result = reviewer.applyDocumentOperationsToStory({
        undefinedReferences: "refuse",
        story,
        batch: {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [
            type === "mergeBlockWithNext"
              ? { id: "note-change", type, blockId: block.id, separator: " " }
              : { id: "note-change", type, blockId: block.id },
          ],
        },
      });
      assert.equal(result.applied.length, 1, JSON.stringify(result.issues));
      await assertAcceptedReaders(
        new Uint8Array(await reviewer.toBuffer()),
        `${story.type} ${type}: referenced structural note`,
      );
    }
    for (const order of ["paragraph-first", "table-first"] as const) {
      const reviewer = await openReviewer(fixture);
      const snapshot = reviewer.snapshotStory(story);
      assert.ok(snapshot);
      const paragraph = snapshot.blocks.at(1);
      const cell = snapshot.blocks.find(({ table }) => table !== undefined);
      assert.ok(paragraph && cell);
      const paragraphDeletion = {
        id: "note-paragraph",
        type: "deleteBlock",
        blockId: paragraph.id,
      } as const satisfies FolioDocumentOperation;
      const tableDeletion = {
        id: "note-table",
        type: "deleteTable",
        blockId: cell.id,
      } as const satisfies FolioDocumentOperation;
      const operations =
        order === "paragraph-first"
          ? [paragraphDeletion, tableDeletion]
          : [tableDeletion, paragraphDeletion];
      for (const operation of operations) {
        const result = reviewer.applyDocumentOperationsToStory({
          undefinedReferences: "refuse",
          story,
          batch: {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "tracked-changes",
            operations: [operation],
          },
        });
        assert.equal(result.applied.length, 1, JSON.stringify(result.issues));
      }
      await assertAcceptedReaders(
        new Uint8Array(await reviewer.toBuffer()),
        `${story.type}: run-in across deleted table, ${order}`,
      );
    }
  }
});
