import { test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const SOURCE = [
  ["10000001", "Opening clause."],
  ["10000002", "Signed in two copies."],
  ["10000003", "Closing clause."],
] as const;

const source = () => {
  const document = createEmptyDocument();
  document.package.document.content = SOURCE.map(([paraId, text]) =>
    Object.assign(paragraph(text), { paraId }),
  );
  return createDocx(document);
};

type Batch = {
  mode: "tracked-changes" | "suggested";
  operation: Record<string, unknown>;
};
const CASES: [string, Batch[], number[]][] = [
  [
    "1530013392",
    [
      {
        mode: "suggested",
        operation: {
          type: "insertAfterBlock",
          blockId: "10000002",
          text: "Payment due.",
        },
      },
      {
        mode: "tracked-changes",
        operation: { type: "deleteBlock", blockId: "10000003" },
      },
      {
        mode: "tracked-changes",
        operation: {
          type: "insertAfterBlock",
          blockId: "10000002",
          text: "Payment due.",
        },
      },
    ],
    [93520625, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  ],
  [
    "501932560",
    [
      {
        mode: "tracked-changes",
        operation: { type: "deleteBlock", blockId: "10000003" },
      },
      {
        mode: "suggested",
        operation: {
          type: "insertAfterBlock",
          blockId: "10000002",
          text: "Payment due.",
        },
      },
      {
        mode: "tracked-changes",
        operation: {
          type: "insertBeforeBlock",
          blockId: "10000003",
          text: "Payment due.",
        },
      },
    ],
    [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  ],
  [
    "-1819669582",
    [
      {
        mode: "suggested",
        operation: {
          type: "insertAfterBlock",
          blockId: "10000003",
          text: "Payment due.",
        },
      },
      {
        mode: "tracked-changes",
        operation: {
          type: "insertAfterBlock",
          blockId: "10000003",
          text: "Payment due.",
        },
      },
      {
        mode: "tracked-changes",
        operation: { type: "splitBlock", blockId: "10000002", offset: 1 },
      },
    ],
    [1414791501, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  ],
  [
    "68121680",
    [
      {
        mode: "tracked-changes",
        operation: { type: "deleteBlock", blockId: "10000002" },
      },
      {
        mode: "suggested",
        operation: {
          type: "insertAfterBlock",
          blockId: "10000001",
          text: "Payment due.",
        },
      },
      {
        mode: "tracked-changes",
        operation: { type: "deleteBlock", blockId: "10000003" },
      },
      {
        mode: "tracked-changes",
        operation: {
          type: "insertBeforeBlock",
          blockId: "10000002",
          text: "Payment due.",
        },
      },
    ],
    [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  ],
];

const KEEP = new Set([
  "paraId",
  "pPrMark",
  "_suggestedInsert",
  "type",
  "text",
  "marks",
  "attrs",
  "content",
  "id",
  "revisionId",
  "kind",
  "suggestionId",
  "provenance",
  "info",
]);
const prune = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(prune);
  if (value && typeof value === "object") {
    const out: Record<string, unknown> = {};
    for (const [key, entry] of Object.entries(value)) {
      if (entry === null || entry === undefined || entry === false) continue;
      if (Array.isArray(entry) && entry.length === 0) continue;
      if (key === "author" || key === "date") continue;
      if (!KEEP.has(key) && typeof entry !== "object") continue;
      out[key] = prune(entry);
    }
    return out;
  }
  return value;
};

const dump = (label: string, reviewer: FolioDocxReviewer) => {
  const state = (
    reviewer as unknown as { state: { doc: { toJSON(): unknown } } }
  ).state;
  console.log(
    `DIAG ${label}\n  changes=${JSON.stringify(reviewer.getChanges().map(({ id, type, text, blockId }) => ({ id, type, text, blockId })))}\n  doc=${JSON.stringify(prune(state.doc.toJSON()))}`,
  );
};

for (const [name, batches, picks] of CASES) {
  test(`diag ${name}`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await source());
    const suggestions: string[] = [];
    dump(`${name} start`, reviewer);
    for (const [index, { mode, operation }] of batches.entries()) {
      const id = `op-${String(index)}`;
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [{ ...operation, id } as FolioDocumentOperation],
        },
        { undefinedReferences: "refuse" },
      );
      if (mode === "suggested" && result.applied.length > 0)
        suggestions.push(id);
      dump(
        `${name} after ${id} ${mode} ${JSON.stringify(operation)} applied=${String(result.applied.length)}`,
        reviewer,
      );
    }
    for (const pick of picks) {
      const changes = reviewer.getChanges();
      if (changes.length === 0 && suggestions.length === 0) break;
      if (pick % 2 === 0 && suggestions.length > 0) {
        const [id] = suggestions.splice(pick % suggestions.length, 1);
        const ok = id !== undefined && reviewer.rejectSuggestion(id);
        dump(
          `${name} rejectSuggestion ${String(id)} -> ${String(ok)}`,
          reviewer,
        );
        continue;
      }
      const change = changes[pick % Math.max(changes.length, 1)];
      if (change !== undefined) {
        const ok = reviewer.rejectChange(change);
        dump(
          `${name} rejectChange ${String(change.id)} ${change.type} -> ${String(ok)}`,
          reviewer,
        );
      }
    }
    for (const id of suggestions) reviewer.rejectSuggestion(id);
    reviewer.rejectAll();
    dump(`${name} final`, reviewer);
  });
}
