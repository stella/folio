import type { FolioAIEditOperation, FolioAIResolvedEditOperation } from "../../src/ai-edits/types";
import { resolveNewListOperations } from "../../src/ai-edits/newListNumbering";

const request = {
  id: "numbering-request",
  type: "setBlockParagraphProperties",
  blockId: "paragraph",
  properties: { numbering: { kind: "newList", format: "numbered" } },
} as const satisfies FolioAIEditOperation;

// @ts-expect-error the allocator must resolve newList before application
const unresolved: FolioAIResolvedEditOperation = request;

const resolved = resolveNewListOperations([request], null).operations;
const applied: readonly FolioAIResolvedEditOperation[] = resolved;

export type ResolvedNumberingProof = typeof applied | typeof unresolved;
