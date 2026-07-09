import { describe, expect, test } from "bun:test";

import {
  assertSupportedFolioDocumentOperationVersion,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  getFolioDocumentOperationCapabilities,
  isSupportedFolioDocumentOperationVersion,
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
});
