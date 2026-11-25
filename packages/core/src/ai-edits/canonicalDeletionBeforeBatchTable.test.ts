import { panic } from "better-result";
import { afterAll, beforeAll, expect, setSystemTime, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";
import { hashFolioAIBlockContent } from "./snapshot";

beforeAll(() => setSystemTime(new Date("2026-01-01T12:00:00Z")));
afterAll(() => setSystemTime());

const anchor = "10000000";
const OPERATIONS: readonly FolioDocumentOperation[] = [
  { id: "table", type: "insertTable", blockId: anchor, rows: [["Term"]] },
  { id: "before", type: "insertBeforeBlock", blockId: anchor, text: "Before." },
  { id: "delete", type: "deleteBlock", blockId: anchor },
];
const PERMUTATIONS = [
  [0, 1, 2],
  [0, 2, 1],
  [1, 0, 2],
  [1, 2, 0],
  [2, 0, 1],
  [2, 1, 0],
];

const tracked = (
  reviewer: FolioDocxReviewer,
  operations: readonly FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [...operations],
    },
    { undefinedReferences: "refuse" },
  );

/** A paragraph whose own mark a tracked split inserted, holding original words. */
const splitDocument = async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await paragraphsDocx([[textRun("Alpha bravo.")]]),
    { session: "canonical" },
  );
  tracked(reviewer, [
    { id: "split", type: "splitBlock", blockId: anchor, offset: 5 },
  ]);
  return reviewer;
};

// Before a table the batch puts after it, the paragraph's deletion takes its
// words only, as it does once that table stands: tracked, in any order.
test("a deletion before a table its batch inserts reads as after that table stands", async () => {
  const sequential = await splitDocument();
  for (const operation of OPERATIONS)
    expect(tracked(sequential, [operation]).skipped).toEqual([]);
  sequential.acceptAll();
  const expected = sequential.getContent().map(({ text }) => text);
  for (const order of PERMUTATIONS) {
    const reviewer = await splitDocument();
    const result = tracked(
      reviewer,
      order.flatMap((index) => OPERATIONS[index] ?? []),
    );
    expect(result.skipped).toEqual([]);
    reviewer.acceptAll();
    expect(reviewer.getContent().map(({ text }) => text)).toEqual(expected);
  }
});

for (const before of [
  "survives",
  "removed",
  "reveals-table",
  "refused",
] as const) {
  for (const after of ["survives", "removed", "reveals-table"] as const) {
    test(`a deleted separator with ${before} before it and ${after} after it`, async () => {
      const table = (label: string) => `| ${label} |\n| --- |`;
      const content = [
        "Opening.",
        ...(before === "reveals-table" ? [table("Outer before")] : []),
        table("Inner before"),
        "Separator.",
        table("Inner after"),
        ...(after === "reveals-table" ? [table("Outer after")] : []),
        "Closing.",
      ].join("\n\n");
      const bytes = (
        await prepareCanonicalDocxInput(await createDocx(fromMarkdown(content)))
      )
        .unwrap()
        .slice().buffer;
      const reviewer = await FolioDocxReviewer.fromBuffer(bytes);
      const blockId = (text: string) =>
        reviewer.getContent().find((block) => block.text === text)?.id ??
        panic("The separator fixture needs each original block.");
      const toDelete = [
        ...(before === "survives" ? [] : ["Inner before"]),
        ...(after === "survives" ? [] : ["Inner after"]),
      ];
      const refused = before === "refused" ? ["Inner before"] : [];
      const deleteTables = toDelete.map((text) => {
        const deletion: Extract<
          FolioDocumentOperation,
          { type: "deleteTable" }
        > = { id: text, type: "deleteTable", blockId: blockId(text) };
        if (refused.includes(text))
          deletion.precondition = {
            blockTextHash: hashFolioAIBlockContent("Stale."),
          };
        return deletion;
      });
      if (deleteTables.length > 0) {
        const result = tracked(reviewer, deleteTables);
        expect(result.applied).toHaveLength(
          deleteTables.length - refused.length,
        );
        expect(result.skipped.map(({ id }) => id)).toEqual(refused);
      }
      const separator = {
        id: "separator",
        type: "deleteBlock",
        blockId: blockId("Separator."),
      } as const satisfies FolioDocumentOperation;
      const deletion = tracked(reviewer, [separator]);
      expect(deletion.applied).toHaveLength(1);
      expect(deletion.skipped).toEqual([]);
      expect(reviewer.acceptAll().skipped).toEqual([]);
      const expected = [
        "Opening.",
        ...(before === "removed"
          ? []
          : [before === "reveals-table" ? "Outer before" : "Inner before"]),
        ...(before === "removed" || after === "removed" ? [] : [""]),
        ...(after === "removed"
          ? []
          : [after === "survives" ? "Inner after" : "Outer after"]),
        "Closing.",
      ];
      const texts = (opened: FolioDocxReviewer) =>
        opened.getContent().map(({ text }) => text);
      expect(texts(reviewer)).toEqual(expected);

      const operations = [...deleteTables, separator];
      for (const mode of ["direct", "tracked-changes"] as const) {
        const batched = await FolioDocxReviewer.fromBuffer(bytes);
        const result = batched.applyDocumentOperations(
          {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode,
            operations,
          },
          { undefinedReferences: "refuse" },
        );
        expect(result.skipped.map(({ id }) => id)).toEqual(refused);
        expect(result.applied).toHaveLength(operations.length - refused.length);
        if (mode === "tracked-changes")
          expect(batched.acceptAll().skipped).toEqual([]);
        expect(texts(batched)).toEqual(expected);
      }
    });
  }
}
