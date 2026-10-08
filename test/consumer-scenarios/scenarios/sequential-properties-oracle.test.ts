import assert from "node:assert/strict";
import { test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "@stll/folio-core/server";

import { openReviewer, packDocument } from "../support/documents.ts";
import { sequentialGroups } from "../support/metamorphic.ts";
import { type Row } from "../support/oracle.ts";

const paragraphRows = (count: number): Row[] =>
  Array.from({ length: count }, (_, index) => ({
    id: `paragraph-${index}`,
    text: `Clause ${index}.`,
    kind: "paragraph",
  }));

test("sequential replay orders every property-edit sibling before terminal break retirement", () => {
  for (let count = 2; count <= 6; count++) {
    const preRows = paragraphRows(count);
    for (let removed = 1; removed < count; removed++) {
      const carrier = preRows.at(count - removed - 1);
      assert.ok(carrier);
      const deletions = preRows.slice(count - removed).map(({ id }) => ({
        id: `delete-${id}`,
        type: "deleteBlock",
        blockId: id,
      }));
      for (const property of [
        {
          id: "patch",
          type: "setBlockParagraphProperties",
          blockId: carrier.id,
          properties: { spacing: null },
        },
        {
          id: "restyle",
          type: "replaceBlock",
          blockId: carrier.id,
          text: carrier.text,
          styleId: "Heading2",
        },
      ]) {
        for (const applied of [
          [property, ...deletions],
          [...deletions, property],
        ]) {
          for (const mode of ["direct", "suggested", "tracked-changes"] as const) {
            const ordered = sequentialGroups({ applied, preRows, mode }).flat();
            assert.deepEqual(
              ordered.map(({ id }) => id).toSorted(),
              applied.map(({ id }) => id).toSorted(),
            );
            assert.equal(ordered.at(mode === "tracked-changes" ? 0 : -1)?.id, property.id);
          }
        }
      }
    }
  }
});

test("appending after the old last paragraph leaves ordinary replay ordering", () => {
  const preRows = paragraphRows(3);
  const applied = [
    {
      id: "patch",
      type: "setBlockParagraphProperties",
      blockId: "paragraph-1",
      properties: { alignment: "center" },
    },
    { id: "delete", type: "deleteBlock", blockId: "paragraph-2" },
    { id: "append", type: "insertAfterBlock", blockId: "paragraph-2", text: "New last paragraph." },
  ];
  assert.deepEqual(
    sequentialGroups({ applied, preRows, mode: "tracked-changes" })
      .flat()
      .map(({ id }) => id),
    ["append", "delete", "patch"],
  );
});

test("terminal retirement ordering stays within each table cell", () => {
  const preRows = paragraphRows(4).map((row, index) =>
    Object.assign(row, {
      table: {
        outerTableIndex: 0,
        tableIndex: 0,
        rowIndex: 0,
        cellIndex: Math.floor(index / 2),
        gridColumnIndex: Math.floor(index / 2),
        columnSpan: 1,
        rowSpan: 1,
      },
    }),
  );
  const applied = [
    {
      id: "left-patch",
      type: "setBlockParagraphProperties",
      blockId: "paragraph-0",
      properties: { alignment: "center" },
    },
    {
      id: "right-patch",
      type: "setBlockParagraphProperties",
      blockId: "paragraph-2",
      properties: { alignment: "center" },
    },
    { id: "left-delete", type: "deleteBlock", blockId: "paragraph-1" },
  ];
  assert.deepEqual(
    sequentialGroups({ applied, preRows, mode: "tracked-changes" })
      .flat()
      .map(({ id }) => id),
    ["left-patch", "right-patch", "left-delete"],
  );
  assert.throws(
    () =>
      sequentialGroups({
        applied: [{ id: "unknown", type: "unknown" }],
        preRows,
        mode: "tracked-changes",
      }),
    /no policy for unknown/u,
  );
});

test("generated terminal deletions preserve property edits and rejection in real sequential replay", async () => {
  for (let removed = 1; removed <= 3; removed++) {
    for (const properties of [
      { styleId: "Heading2" },
      { alignment: "center" },
      { spacing: { spaceBefore: 120, spaceAfter: 120 } },
      { spacing: null },
    ] as const) {
      const fixture = await packDocument(
        fromMarkdown(
          Array.from({ length: removed + 2 }, (_, index) => `Clause ${index}.`).join("\n\n"),
        ),
      );
      const batch = await openReviewer(fixture);
      const originalRows = batch.getContent();
      const carrier = originalRows.at(1);
      assert.ok(carrier);
      const initialization = {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [
          {
            id: "original-spacing",
            type: "setBlockParagraphProperties",
            blockId: carrier.id,
            properties: { spacing: { spaceBefore: 240, spaceAfter: 240 } },
          },
        ],
      } as const satisfies Parameters<typeof batch.applyDocumentOperations>[0];
      assert.equal(batch.applyDocumentOperations(initialization).applied.length, 1);
      const before = new Uint8Array(await batch.toBuffer());
      const sequential = await openReviewer(before);
      const rejected = await openReviewer(before);
      const preRows = batch.getContent();
      const operations = [
        { id: "properties", type: "setBlockParagraphProperties", blockId: carrier.id, properties },
        ...originalRows
          .slice(2)
          .map(({ id }) => ({ id: `delete-${id}`, type: "deleteBlock" as const, blockId: id })),
      ] satisfies FolioDocumentOperation[];
      const request = {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations,
      } as const satisfies Parameters<typeof batch.applyDocumentOperations>[0];
      assert.equal(batch.applyDocumentOperations(request).applied.length, operations.length);
      assert.equal(rejected.applyDocumentOperations(request).applied.length, operations.length);
      for (const group of sequentialGroups({ applied: operations, preRows, mode: request.mode })) {
        const restated = group.map(({ id }) => {
          const operation = operations.find((candidate) => candidate.id === id);
          assert.ok(operation, `sequential group operation ${id} belongs to the batch`);
          return operation;
        });
        const result = sequential.applyDocumentOperations({ ...request, operations: restated });
        assert.deepEqual(
          result.skipped,
          [],
          "every applied operation must re-resolve sequentially",
        );
      }
      batch.acceptAll();
      sequential.acceptAll();
      assert.deepEqual(sequential.getContent(), batch.getContent());
      rejected.rejectAll();
      assert.deepEqual(rejected.getContent(), preRows);
    }
  }
});
