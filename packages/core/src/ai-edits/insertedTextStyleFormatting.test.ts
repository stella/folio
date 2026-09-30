/**
 * An inserted paragraph's words take their run formatting from the
 * paragraph's style. Text the operation gave no emphasis states no run
 * property of its own, so a heading's words stay as bold as its style makes
 * them, before and after a save.
 */

import { describe, expect, test } from "bun:test";

import { createDocx } from "../docx/rezip";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const MODES = [
  "direct",
  "tracked-changes",
] as const satisfies readonly FolioDocumentOperationMode[];
const INSERTED = "Inserted heading.";

const insertedRuns = async (
  mode: FolioDocumentOperationMode,
  styleId: string | undefined,
  text = INSERTED,
) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(fromMarkdown("# Agreement\n\n## Scope\n\nBody text.")),
    { author: "AI" },
  );
  const blockId = reviewer.getContent().find((block) => block.text === "Agreement")?.id;
  const result = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: [
      { id: "op", type: "insertAfterBlock", blockId, text, ...(styleId && { styleId }) } as never,
    ],
  });
  expect(result.skipped).toEqual([]);
  if (mode !== "direct") reviewer.acceptAll();
  const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
  return [reviewer, reopened].map(
    (current) =>
      current.getContent().find((block) => block.text.startsWith("Inserted"))?.previewRuns ?? [],
  );
};

describe("an inserted paragraph's words follow its style", () => {
  for (const mode of MODES) {
    for (const styleId of ["Heading2", undefined]) {
      test(`after a heading, style ${String(styleId)} (${mode})`, async () => {
        for (const runs of await insertedRuns(mode, styleId)) {
          expect(runs.length).toBeGreaterThan(0);
          for (const run of runs) {
            expect(run.bold).toBe(true);
            expect(run.directFormatting).toBeUndefined();
          }
        }
      });
    }

    test(`emphasis the operation asks for stays its own (${mode})`, async () => {
      for (const runs of await insertedRuns(mode, "Heading2", "Inserted ***heading***.")) {
        expect(runs.find((run) => run.text === "heading")?.italic).toBe(true);
        expect(runs.find((run) => run.text === "Inserted ")?.directFormatting).toBeUndefined();
      }
    });
  }
});
