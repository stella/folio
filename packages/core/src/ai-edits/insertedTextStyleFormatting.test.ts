/**
 * An inserted paragraph's words take their run formatting from the
 * paragraph's style. Text the operation gave no emphasis states no run
 * property of its own, so a heading's words stay as bold as its style makes
 * them, before and after a save.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

const MODES = [
  "direct",
  "tracked-changes",
] as const satisfies readonly FolioDocumentOperationMode[];
const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];
const INSERTED = "Inserted heading.";

type InsertedRunsOptions = {
  session: HeadlessReviewerSession;
  mode: FolioDocumentOperationMode;
  styleId: string | undefined;
  text?: string;
};

const insertedRuns = async ({
  session,
  mode,
  styleId,
  text = INSERTED,
}: InsertedRunsOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(fromMarkdown("# Agreement\n\n## Scope\n\nBody text.")),
    { author: "AI", session },
  );
  const blockId = reviewer
    .getContent()
    .find((block) => block.text === "Agreement")?.id;
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [
        {
          id: "op",
          type: "insertAfterBlock",
          blockId,
          text,
          ...(styleId && { styleId }),
        } as never,
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  if (mode !== "direct") reviewer.acceptAll();
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
    { session },
  );
  return [reviewer, reopened].map(
    (current) =>
      current.getContent().find((block) => block.text.startsWith("Inserted"))
        ?.previewRuns ?? [],
  );
};

describe.each(SESSIONS)(
  "an inserted paragraph's words follow its style (%s)",
  (session) => {
    for (const mode of MODES) {
      for (const styleId of ["Heading2", undefined]) {
        test(`after a heading, style ${String(styleId)} (${mode})`, async () => {
          for (const runs of await insertedRuns({ session, mode, styleId })) {
            expect(runs.length).toBeGreaterThan(0);
            for (const run of runs) {
              expect(run.bold).toBe(true);
              expect(run.directFormatting).toBeUndefined();
            }
          }
        });
      }

      test(`emphasis the operation asks for stays its own (${mode})`, async () => {
        for (const runs of await insertedRuns({
          session,
          mode,
          styleId: "Heading2",
          text: "Inserted ***heading***.",
        })) {
          expect(runs.find((run) => run.text === "heading")?.italic).toBe(true);
          expect(
            runs.find((run) => run.text === "Inserted ")?.directFormatting,
          ).toBeUndefined();
        }
      });
    }
  },
);
