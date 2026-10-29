/**
 * Temporary diagnostic: what a tracked merge that restyles the merged
 * paragraph writes, and what each rejection reads.
 */

import { test } from "node:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  parseDocx,
} from "@stll/folio-core/server";

import {
  openReviewer,
  storiesDocument,
  toArrayBuffer,
} from "../support/documents.ts";

const IDS = new Set(["7735D914", "7B806329"]);

const pick = (blocks: readonly unknown[]): unknown[] => {
  const found: unknown[] = [];
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    if (
      "type" in value &&
      value.type === "paragraph" &&
      "paraId" in value &&
      typeof value.paraId === "string" &&
      IDS.has(value.paraId)
    ) {
      found.push(value);
      return;
    }
    for (const child of Object.values(value)) visit(child);
  };
  visit(blocks);
  return found;
};

test("diagnose merged restyle", async () => {
  const live = await openReviewer(await storiesDocument());
  const result = live.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: "7735D914",
          separator: " ",
          mergedParagraphProperties: { styleId: "Heading2" },
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  console.log("DIAG applied", JSON.stringify(result.applied));
  const bytes = new Uint8Array(await live.toBuffer());
  const parsed = await parseDocx(toArrayBuffer(bytes), { preloadFonts: false });
  console.log(
    "DIAG saved",
    JSON.stringify(pick(parsed.package.document.content), null, 1),
  );
  const reopened = await openReviewer(bytes);
  for (const decision of ["reject", "accept"] as const) {
    console.log(
      `DIAG live ${decision}`,
      JSON.stringify(
        pick(live.toResolvedDocument(decision).package.document.content),
        null,
        1,
      ),
    );
    console.log(
      `DIAG reopened ${decision}`,
      JSON.stringify(
        pick(reopened.toResolvedDocument(decision).package.document.content),
        null,
        1,
      ),
    );
  }
  console.log("DIAG live changes", JSON.stringify(live.getChanges()));
  console.log("DIAG reopened changes", JSON.stringify(reopened.getChanges()));
});
