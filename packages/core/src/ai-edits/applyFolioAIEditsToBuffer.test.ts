/**
 * The one-shot review helper on both engines: operations built against a
 * snapshot of the same buffer, from either engine, apply the same way.
 */
import { expect, test } from "bun:test";

import { textRun } from "../__tests__/operationBatchDocuments";
import { createDocx } from "../docx/rezip";
import { unzipDocx } from "../docx/unzip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { applyFolioAIEditsToBuffer, FolioDocxReviewer } from "./headless";

const SESSIONS = ["legacy", "canonical"] as const;

const source = () =>
  createDocx({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "10000001",
            content: [textRun("Heading text")],
          },
          {
            type: "paragraph",
            paraId: "10000002",
            content: [textRun("Body text")],
          },
        ],
      },
    },
  });

const run = async (
  session: HeadlessReviewerSession,
  snapshotSession: HeadlessReviewerSession,
  mode: "direct" | "tracked-changes",
) => {
  const buffer = await source();
  const snapshot = (
    await FolioDocxReviewer.fromBuffer(buffer, { session: snapshotSession })
  ).snapshot();
  const target = snapshot.blocks.find(
    (block) => block.kind !== "diagnostic" && block.text.startsWith("Heading"),
  );
  if (target === undefined) throw new Error("no heading block");
  return applyFolioAIEditsToBuffer(
    buffer,
    [
      {
        id: "o1",
        type: "replaceInBlock",
        blockId: target.id,
        find: "Heading",
        replace: "Intro",
      },
      { id: "o2", type: "insertAfterBlock", blockId: target.id, text: "Added" },
    ],
    { mode, snapshot, session, author: "Agent" },
  );
};

for (const session of SESSIONS)
  for (const snapshotSession of SESSIONS)
    for (const mode of ["direct", "tracked-changes"] as const)
      test(`${session} applies operations built on a ${snapshotSession} snapshot (${mode})`, async () => {
        const result = await run(session, snapshotSession, mode);
        expect(result.skipped).toEqual([]);
        expect(result.applied.map(({ id }) => id).toSorted()).toEqual([
          "o1",
          "o2",
        ]);
        const reopened = await FolioDocxReviewer.fromBuffer(result.buffer, {
          session,
        });
        const accepted =
          mode === "direct"
            ? reopened.toDocument()
            : reopened.toResolvedDocument("accept");
        const xml = (await unzipDocx(result.buffer)).documentXml ?? "";
        expect(xml).toContain("Intro");
        expect(JSON.stringify(accepted.package.document.content)).toContain(
          "Added",
        );
        if (mode === "tracked-changes") expect(xml).toContain("<w:ins ");
      });
