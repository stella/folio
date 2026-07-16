import { describe, expect, test } from "bun:test";

import {
  assertSupportedFolioDocumentOperationVersion,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_DOCUMENT_OPERATION_BATCH_MODES,
  FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE,
  FOLIO_DOCUMENT_OPERATION_PRECONDITIONS,
  getFolioDocumentOperationCapabilities,
  getFolioDocumentOperationReceipts,
  isFolioDocumentOperationModeSupported,
  InvalidFolioDocumentOperationBatchError,
  isSupportedFolioDocumentOperationVersion,
  parseFolioDocumentOperationBatch,
  UnsupportedFolioDocumentOperationVersionError,
  type FolioDocumentOperation,
} from "./document-operations";

describe("document operation contract", () => {
  test("reports the supported version, operations, modes, and stories", () => {
    expect(getFolioDocumentOperationCapabilities()).toEqual({
      version: 1,
      operationTypes: [
        "replaceInBlock",
        "replaceRange",
        "commentOnRange",
        "formatRange",
        "insertAfterBlock",
        "insertBeforeBlock",
        "replaceBlock",
        "deleteBlock",
        "commentOnBlock",
        "insertSignatureTable",
        "insertTableRow",
        "deleteTableRow",
        "insertTableColumn",
        "deleteTableColumn",
        "mergeTableCells",
        "splitTableCell",
      ],
      modes: ["direct", "tracked-changes"],
      batchModes: ["best-effort", "atomic"],
      dryRun: true,
      modesByOperationType: {
        replaceInBlock: ["direct", "tracked-changes"],
        replaceRange: ["direct", "tracked-changes"],
        commentOnRange: ["direct", "tracked-changes"],
        formatRange: ["direct"],
        insertAfterBlock: ["direct", "tracked-changes"],
        insertBeforeBlock: ["direct", "tracked-changes"],
        replaceBlock: ["direct", "tracked-changes"],
        deleteBlock: ["direct", "tracked-changes"],
        commentOnBlock: ["direct", "tracked-changes"],
        insertSignatureTable: ["direct"],
        insertTableRow: ["direct"],
        deleteTableRow: ["direct"],
        insertTableColumn: ["direct"],
        deleteTableColumn: ["direct"],
        mergeTableCells: ["direct"],
        splitTableCell: ["direct"],
      },
      preconditions: ["blockTextHash"],
      stories: ["main", "header", "footer", "footnote", "endnote"],
    });
  });

  test("reports mode support for each operation type", () => {
    expect(isFolioDocumentOperationModeSupported("replaceInBlock", "tracked-changes")).toBe(true);
    expect(isFolioDocumentOperationModeSupported("insertSignatureTable", "direct")).toBe(true);
    expect(isFolioDocumentOperationModeSupported("insertSignatureTable", "tracked-changes")).toBe(
      false,
    );
    expect(isFolioDocumentOperationModeSupported("deleteTableRow", "tracked-changes")).toBe(false);
    expect(isFolioDocumentOperationModeSupported("insertTableColumn", "direct")).toBe(true);
    expect(isFolioDocumentOperationModeSupported("deleteTableColumn", "direct")).toBe(true);
    expect(isFolioDocumentOperationModeSupported("mergeTableCells", "direct")).toBe(true);
    expect(isFolioDocumentOperationModeSupported("splitTableCell", "direct")).toBe(true);
    expect(
      Reflect.apply(isFolioDocumentOperationModeSupported, null, ["unknownOperation", "direct"]),
    ).toBe(false);
    expect(Object.keys(FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE)).toEqual(
      getFolioDocumentOperationCapabilities().operationTypes,
    );
    expect(FOLIO_DOCUMENT_OPERATION_PRECONDITIONS).toEqual(["blockTextHash"]);
    expect(FOLIO_DOCUMENT_OPERATION_BATCH_MODES).toEqual(["best-effort", "atomic"]);
  });

  test("checks untyped contract versions at a serialization boundary", () => {
    expect(isSupportedFolioDocumentOperationVersion(1)).toBe(true);
    expect(isSupportedFolioDocumentOperationVersion("1")).toBe(false);
    expect(isSupportedFolioDocumentOperationVersion(2)).toBe(false);
    expect(FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION).toBe(1);
  });

  test("builds input-ordered receipts for successful affected targets", () => {
    const range = {
      type: "textRange",
      story: "main",
      blockId: "paragraph-2",
      startOffset: 4,
      endOffset: 9,
      selectedTextHash: "h123",
    } as const;
    const operations = [
      {
        id: "replace",
        type: "replaceInBlock",
        blockId: "paragraph-1",
        find: "old",
        replace: "new",
        comment: { text: "Changed" },
      },
      { id: "comment", type: "commentOnRange", range, comment: { text: "Review" } },
      { id: "format", type: "formatRange", range, formatting: { bold: true } },
      { id: "insert", type: "insertBeforeBlock", blockId: "paragraph-3", text: "New" },
      { id: "delete", type: "deleteBlock", blockId: "paragraph-4" },
      {
        id: "signature",
        type: "insertSignatureTable",
        blockId: "paragraph-5",
        parties: [{ name: "Party" }],
      },
      {
        id: "row",
        type: "insertTableRow",
        blockId: "paragraph-6",
        position: "before",
        cellTexts: ["A", "B"],
      },
      { id: "delete-row", type: "deleteTableRow", blockId: "paragraph-7" },
      {
        id: "column",
        type: "insertTableColumn",
        blockId: "paragraph-8",
        cellTexts: ["A", "B"],
      },
      { id: "delete-column", type: "deleteTableColumn", blockId: "paragraph-9" },
      {
        id: "merge-cells",
        type: "mergeTableCells",
        blockId: "paragraph-10",
        endBlockId: "paragraph-11",
      },
      { id: "split-cell", type: "splitTableCell", blockId: "paragraph-12" },
    ] as const satisfies readonly FolioDocumentOperation[];

    expect(
      getFolioDocumentOperationReceipts(operations, [
        { id: "signature" },
        { id: "format" },
        { id: "replace", commentId: 17 },
        { id: "insert" },
        { id: "comment", commentId: 18 },
        { id: "delete" },
        { id: "row" },
        { id: "delete-row" },
        { id: "column" },
        { id: "delete-column" },
        { id: "merge-cells" },
        { id: "split-cell" },
      ]),
    ).toEqual([
      {
        operationId: "replace",
        operationIndex: 0,
        affected: [
          {
            type: "block",
            story: "main",
            blockId: "paragraph-1",
            effect: "updated",
          },
          { type: "comment", commentId: 17 },
        ],
      },
      {
        operationId: "comment",
        operationIndex: 1,
        affected: [
          { type: "textRange", range, effect: "commented" },
          { type: "comment", commentId: 18 },
        ],
      },
      {
        operationId: "format",
        operationIndex: 2,
        affected: [{ type: "textRange", range, effect: "formatted" }],
      },
      {
        operationId: "insert",
        operationIndex: 3,
        affected: [
          {
            type: "insertion",
            story: "main",
            anchorBlockId: "paragraph-3",
            position: "before",
            content: "block",
          },
        ],
      },
      {
        operationId: "delete",
        operationIndex: 4,
        affected: [
          {
            type: "block",
            story: "main",
            blockId: "paragraph-4",
            effect: "deleted",
          },
        ],
      },
      {
        operationId: "signature",
        operationIndex: 5,
        affected: [
          {
            type: "insertion",
            story: "main",
            anchorBlockId: "paragraph-5",
            position: "after",
            content: "signatureTable",
          },
        ],
      },
      {
        operationId: "row",
        operationIndex: 6,
        affected: [
          {
            type: "insertion",
            story: "main",
            anchorBlockId: "paragraph-6",
            position: "before",
            content: "tableRow",
          },
        ],
      },
      {
        operationId: "delete-row",
        operationIndex: 7,
        affected: [
          {
            type: "tableRow",
            story: "main",
            anchorBlockId: "paragraph-7",
            effect: "deleted",
          },
        ],
      },
      {
        operationId: "column",
        operationIndex: 8,
        affected: [
          {
            type: "insertion",
            story: "main",
            anchorBlockId: "paragraph-8",
            position: "after",
            content: "tableColumn",
          },
        ],
      },
      {
        operationId: "delete-column",
        operationIndex: 9,
        affected: [
          {
            type: "tableColumn",
            story: "main",
            anchorBlockId: "paragraph-9",
            effect: "deleted",
          },
        ],
      },
      {
        operationId: "merge-cells",
        operationIndex: 10,
        affected: [
          {
            type: "tableCells",
            story: "main",
            anchorBlockId: "paragraph-10",
            endAnchorBlockId: "paragraph-11",
            effect: "merged",
          },
        ],
      },
      {
        operationId: "split-cell",
        operationIndex: 11,
        affected: [
          {
            type: "tableCell",
            story: "main",
            anchorBlockId: "paragraph-12",
            effect: "split",
          },
        ],
      },
    ]);
  });

  test("rejects an unsupported serialized contract version", () => {
    expect(() => assertSupportedFolioDocumentOperationVersion(2)).toThrow(
      UnsupportedFolioDocumentOperationVersionError,
    );
  });

  test("parses every v1 operation variant from serialized input", () => {
    const batch = {
      version: 1,
      mode: "tracked-changes",
      atomic: true,
      dryRun: true,
      operations: [
        {
          id: "1",
          type: "replaceInBlock",
          blockId: "a",
          find: "old",
          replace: "new",
          precondition: { blockTextHash: "h123" },
        },
        {
          id: "range",
          type: "replaceRange",
          range: {
            type: "textRange",
            story: "main",
            blockId: "a",
            startOffset: 0,
            endOffset: 3,
            selectedTextHash: "h123",
          },
          replace: "new",
        },
        { id: "2", type: "insertAfterBlock", blockId: "a", text: "after" },
        { id: "3", type: "insertBeforeBlock", blockId: "a", text: "before" },
        { id: "4", type: "replaceBlock", blockId: "a", text: "replacement" },
        { id: "5", type: "deleteBlock", blockId: "a" },
        {
          id: "6",
          type: "commentOnBlock",
          blockId: "a",
          quote: "text",
          comment: { text: "note" },
        },
        {
          id: "7",
          type: "insertSignatureTable",
          blockId: "a",
          position: "before",
          parties: [{ name: "Party", signatory: "Signer", title: "Director" }],
          severity: "high",
          area: "Execution",
        },
        {
          id: "8",
          type: "insertTableRow",
          blockId: "a",
          position: "after",
          cellTexts: ["First", "Second"],
        },
        { id: "9", type: "deleteTableRow", blockId: "a" },
        {
          id: "10",
          type: "insertTableColumn",
          blockId: "a",
          position: "before",
          cellTexts: ["Top", "Bottom"],
        },
        { id: "11", type: "deleteTableColumn", blockId: "a" },
        {
          id: "12",
          type: "mergeTableCells",
          blockId: "a",
          endBlockId: "b",
        },
        { id: "13", type: "splitTableCell", blockId: "a" },
      ],
    };

    expect(parseFolioDocumentOperationBatch(batch)).toEqual(batch);
  });

  const malformedBatches = [
    [null, "$", "expected an object"],
    [{ version: 1 }, "$.operations", "expected an array"],
    [{ version: 1, operations: [], mode: "review" }, "$.mode", "expected"],
    [{ version: 1, operations: [], atomic: "yes" }, "$.atomic", "expected a boolean"],
    [{ version: 1, operations: [], dryRun: "yes" }, "$.dryRun", "expected a boolean"],
    [
      {
        version: 1,
        operations: [
          { id: "duplicate", type: "deleteBlock", blockId: "a" },
          { id: "duplicate", type: "deleteBlock", blockId: "b" },
        ],
      },
      "$.operations[1].id",
      "expected a unique operation id",
    ],
    [
      {
        version: 1,
        operations: [
          {
            id: "1",
            type: "deleteBlock",
            blockId: "a",
            precondition: { blockTextHash: "not-a-hash" },
          },
        ],
      },
      "$.operations[0].precondition.blockTextHash",
      "expected a normalized block text hash",
    ],
    [
      { version: 1, operations: [{ id: "1", type: "unknown", blockId: "a" }] },
      "$.operations[0].type",
      "unsupported operation type",
    ],
    [
      { version: 1, operations: [{ id: "1", type: "replaceInBlock", blockId: "a" }] },
      "$.operations[0].find",
      "expected a string",
    ],
    [
      {
        version: 1,
        operations: [
          {
            id: "1",
            type: "replaceRange",
            range: {
              type: "textRange",
              story: "main",
              blockId: "a",
              startOffset: 3,
              endOffset: 3,
              selectedTextHash: "h123",
            },
            replace: "new",
          },
        ],
      },
      "$.operations[0].range.endOffset",
      "greater than startOffset",
    ],
    [
      {
        version: 1,
        operations: [{ id: "1", type: "deleteBlock", blockId: "a", typo: true }],
      },
      "$.operations[0].typo",
      "unexpected property",
    ],
    [
      {
        version: 1,
        operations: [{ id: "1", type: "commentOnBlock", blockId: "a", comment: "note" }],
      },
      "$.operations[0].comment",
      "expected an object",
    ],
    [
      {
        version: 1,
        operations: [
          {
            id: "1",
            type: "insertSignatureTable",
            blockId: "a",
            parties: [{ name: 42 }],
          },
        ],
      },
      "$.operations[0].parties[0].name",
      "expected a string",
    ],
    [
      {
        version: 1,
        operations: [
          {
            id: "1",
            type: "insertTableRow",
            blockId: "a",
            cellTexts: ["valid", 42],
          },
        ],
      },
      "$.operations[0].cellTexts[1]",
      "expected a string",
    ],
  ] as const;

  for (const [value, path, reason] of malformedBatches) {
    test(`rejects malformed serialized input at ${path}`, () => {
      expect(() => parseFolioDocumentOperationBatch(value)).toThrow(
        InvalidFolioDocumentOperationBatchError,
      );
      expect(() => parseFolioDocumentOperationBatch(value)).toThrow(path);
      expect(() => parseFolioDocumentOperationBatch(value)).toThrow(reason);
    });
  }
});
