import { describe, expect, test } from "bun:test";

import {
  assertSupportedFolioDocumentOperationVersion,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  getFolioDocumentOperationCapabilities,
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
        "insertAfterBlock",
        "insertBeforeBlock",
        "replaceBlock",
        "deleteBlock",
        "commentOnBlock",
        "insertSignatureTable",
      ],
      modes: ["direct", "tracked-changes"],
      stories: ["main"],
    });
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
      operations: [
        { id: "1", type: "replaceInBlock", blockId: "a", find: "old", replace: "new" },
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
