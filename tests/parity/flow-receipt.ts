import type {
  FolioDocumentOperationBatch,
  FolioDocumentOperationResult,
} from "../../packages/core/src/server";

export type GeneratedFlowOperation = Extract<
  FolioDocumentOperationBatch["operations"][number],
  { type: "insertAfterBlock" | "replaceInBlock" }
>;

type ReceiptTarget = FolioDocumentOperationResult["receipts"][number]["affected"][number];

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const assertReceiptTarget = (value: unknown, operation: GeneratedFlowOperation): void => {
  if (!isRecord(value)) throw new Error("Receipt target is not an object");

  const expected = (() => {
    switch (operation.type) {
      case "insertAfterBlock":
        return {
          type: "insertion",
          story: "main",
          anchorBlockId: operation.blockId,
          position: "after",
          content: "block",
        } as const satisfies ReceiptTarget;
      case "replaceInBlock":
        return {
          type: "block",
          story: "main",
          blockId: operation.blockId,
          effect: "updated",
        } as const satisfies ReceiptTarget;
      default: {
        const exhaustive: never = operation;
        throw new Error(`Unhandled generated operation: ${exhaustive}`);
      }
    }
  })();

  if (!Object.entries(expected).every(([key, expectedValue]) => value[key] === expectedValue)) {
    switch (operation.type) {
      case "insertAfterBlock":
        throw new Error("Insert operation receipt does not describe its insertion");
      case "replaceInBlock":
        throw new Error("Replace operation receipt does not describe its block update");
      default: {
        const exhaustive: never = operation;
        throw new Error(`Unhandled generated operation: ${exhaustive}`);
      }
    }
  }
};

// Validate effectiveness separately from equality: four no-op hosts must fail.
export const assertGeneratedFlowReceipt = (value: unknown, operation: GeneratedFlowOperation) => {
  if (!isRecord(value)) throw new Error("Operation result is not an object");
  if (!Array.isArray(value["skipped"]) || value["skipped"].length !== 0) {
    throw new Error("Generated operation was skipped");
  }
  const applied: unknown = value["applied"];
  if (!Array.isArray(applied) || applied.length !== 1) {
    throw new Error("Applied operation does not match generated operation");
  }
  const appliedEntry: unknown = applied.at(0);
  if (!isRecord(appliedEntry) || appliedEntry["id"] !== operation.id) {
    throw new Error("Applied operation does not match generated operation");
  }
  const receipts: unknown = value["receipts"];
  if (!Array.isArray(receipts) || receipts.length !== 1) {
    throw new Error("Expected one operation receipt");
  }
  const receipt: unknown = receipts.at(0);
  if (!isRecord(receipt)) throw new Error("Missing operation receipt");
  if (receipt["operationId"] !== operation.id || receipt["operationIndex"] !== 0) {
    throw new Error("Operation receipt does not match generated operation");
  }
  const affected: unknown = receipt["affected"];
  if (!Array.isArray(affected) || affected.length !== 1) {
    throw new Error("Expected one affected target");
  }
  assertReceiptTarget(affected.at(0), operation);
};
