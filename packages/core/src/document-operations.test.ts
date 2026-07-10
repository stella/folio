import { describe, expect, test } from "bun:test";

import {
  assertSupportedFolioDocumentOperationVersion,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_DOCUMENT_OPERATION_BATCH_MODES,
  FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE,
  FOLIO_DOCUMENT_OPERATION_PRECONDITIONS,
  getFolioDocumentOperationCapabilities,
  isFolioDocumentOperationModeSupported,
  InvalidFolioDocumentOperationBatchError,
  isSupportedFolioDocumentOperationVersion,
  parseFolioDocumentOperationBatch,
  UnsupportedFolioDocumentOperationVersionError,
} from "./document-operations";

describe("document operation contract", () => {
  test("reports the supported version, operations, modes, and stories", () => {
    expect(getFolioDocumentOperationCapabilities()).toEqual({
      version: 1,
      operationTypes: [
        "replaceInBlock",
        "replaceRange",
        "insertAfterBlock",
        "insertBeforeBlock",
        "replaceBlock",
        "deleteBlock",
        "commentOnBlock",
        "insertSignatureTable",
      ],
      modes: ["direct", "tracked-changes"],
      batchModes: ["best-effort", "atomic"],
      dryRun: true,
      modesByOperationType: {
        replaceInBlock: ["direct", "tracked-changes"],
        replaceRange: ["direct", "tracked-changes"],
        insertAfterBlock: ["direct", "tracked-changes"],
        insertBeforeBlock: ["direct", "tracked-changes"],
        replaceBlock: ["direct", "tracked-changes"],
        deleteBlock: ["direct", "tracked-changes"],
        commentOnBlock: ["direct", "tracked-changes"],
        insertSignatureTable: ["direct"],
      },
      preconditions: ["blockTextHash"],
      stories: ["main"],
    });
  });

  test("reports mode support for each operation type", () => {
    expect(isFolioDocumentOperationModeSupported("replaceInBlock", "tracked-changes")).toBe(true);
    expect(isFolioDocumentOperationModeSupported("insertSignatureTable", "direct")).toBe(true);
    expect(isFolioDocumentOperationModeSupported("insertSignatureTable", "tracked-changes")).toBe(
      false,
    );
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
