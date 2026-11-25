import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  getFolioDocumentOperationReceipts,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { FolioDocxReviewer } from "./headless";

const CASES = ["sole", "joined", "before", "after"] as const;

describe("direct final-paragraph identity receipts", () => {
  for (const session of ["canonical"] as const)
    for (const placement of CASES)
      for (const reversed of [false, true])
        test(`${session} ${placement}, reversed batch: ${reversed}`, async () => {
          const texts = placement === "sole" ? ["C"] : ["A", "B", "C"];
          const buffer = await createDocx({
            package: {
              document: {
                content: texts.map((text, index) => ({
                  type: "paragraph",
                  paraId: (index + 1).toString(16).padStart(8, "0"),
                  content: [{ type: "run", content: [{ type: "text", text }] }],
                })),
              },
            },
          });
          const reviewer = await FolioDocxReviewer.fromBuffer(buffer, {
            author: "Reviewer",
            session,
          });
          const before = reviewer.snapshot();
          const target =
            before.blocks.at(-1) ?? panic("The final paragraph exists.");
          const predecessor = before.blocks.at(-2);
          const operations: FolioDocumentOperation[] = [
            {
              id: "delete",
              type: "deleteBlock",
              blockId: target.id,
            },
          ];
          if (placement === "before" || placement === "after")
            operations.push({
              id: "insert",
              type:
                placement === "before"
                  ? "insertBeforeBlock"
                  : "insertAfterBlock",
              blockId: target.id,
              text: "X",
            });
          if (reversed) operations.reverse();
          const batch = {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "direct",
            operations,
          } as const;
          const preview = reviewer.applyDocumentOperations(
            { ...batch, dryRun: true },
            { undefinedReferences: "refuse" },
          );
          expect(preview.status).toBe("previewed");
          expect(preview.skipped).toEqual([]);
          expect(reviewer.snapshot()).toStrictEqual(before);
          const applied = reviewer.applyDocumentOperations(batch, {
            undefinedReferences: "refuse",
          });
          expect(applied.skipped).toEqual([]);
          const remaining = new Set(
            reviewer.snapshot().blocks.map(({ id }) => id),
          );
          expect(remaining.has(target.id)).toBe(placement !== "after");
          const deletion =
            applied.applied.find(({ id }) => id === "delete") ??
            panic("The deletion applied.");
          const identity = deletion.paragraphIdentity;
          if (placement === "joined") {
            if (!predecessor) panic("The joined paragraph has a predecessor.");
            expect(identity).toEqual({
              type: "joined",
              blockId: target.id,
              retiredBlockId: predecessor.id,
            });
            expect(remaining.has(predecessor.id)).toBe(false);
          } else if (placement === "after") {
            expect(identity).toBeUndefined();
            if (predecessor) expect(remaining.has(predecessor.id)).toBe(true);
          } else {
            expect(identity).toEqual({
              type: "retained",
              blockId: target.id,
            });
            if (predecessor) expect(remaining.has(predecessor.id)).toBe(true);
          }
          for (const receipt of applied.receipts)
            for (const affected of receipt.affected)
              if (affected.type === "block")
                expect(remaining.has(affected.blockId)).toBe(
                  affected.effect !== "deleted",
                );
          expect(
            getFolioDocumentOperationReceipts(operations, applied.applied),
          ).toEqual(applied.receipts);
          expect(
            getFolioDocumentOperationReceipts(operations, preview.applied),
          ).toEqual(preview.receipts);
          expect(preview.receipts).toEqual(applied.receipts);
          expect(preview.applied).toEqual(applied.applied);
        });
});
