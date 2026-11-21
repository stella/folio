/**
 * Text an operation adds right after a note reference is body text: it takes
 * the formatting of the body run around the reference, never the reference's
 * own character style or superscript, on both engines, directly and tracked.
 */

import { describe, expect, test } from "bun:test";

import { fromMarkdown } from "../markdown";
import { createDocx, ensureParaIds } from "../server";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Run } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIEditApplyMode } from "./types";

const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];
const MODES = [
  "direct",
  "tracked-changes",
] as const satisfies readonly FolioAIEditApplyMode[];
const ADDED = " and more";

/** "See the note" in italics, then a footnote reference in its own style. */
const buildDocx = async (): Promise<ArrayBuffer> => {
  const model = fromMarkdown("Placeholder.");
  const first = model.package.document.content[0];
  if (first?.type !== "paragraph") throw new Error("no fixture paragraph");
  first.content = [
    {
      type: "run",
      formatting: { italic: true },
      content: [{ type: "text", text: "See the note" }],
    },
    {
      type: "run",
      formatting: { styleId: "FootnoteReference", vertAlign: "superscript" },
      content: [{ type: "footnoteRef", id: 1 }],
    },
  ];
  model.package.footnotes = [
    {
      type: "footnote",
      id: 1,
      noteType: "normal",
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              formatting: {},
              content: [{ type: "text", text: "The note." }],
            },
          ],
        },
      ],
    },
  ];
  const bytes = (await ensureParaIds(new Uint8Array(await createDocx(model))))
    .docx;
  return bytes.slice().buffer;
};

const source = await buildDocx();

/** The run of the reopened first paragraph that holds the added words. */
const addedRun = (reviewer: FolioDocxReviewer): Run | undefined => {
  const first = reviewer.toDocument().package.document.content[0];
  if (first?.type !== "paragraph") return undefined;
  return first.content
    .filter((item): item is Run => item.type === "run")
    .find((run) =>
      run.content.some((item) => item.type === "text" && item.text === ADDED),
    );
};

describe.each(SESSIONS)("%s session", (session) => {
  test.each(MODES)(
    "%s: text added after a note reference takes the body run's formatting",
    async (mode) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(source.slice(0), {
        author: "Reviewer",
        session,
      });
      const blockId = reviewer.snapshot().blocks[0]?.id;
      if (blockId === undefined) throw new Error("no first block");
      const result = reviewer.applyDocumentOperations(
        {
          version: 1,
          mode,
          operations: [
            {
              id: "append",
              type: "replaceInBlock",
              blockId,
              find: "note[^1]",
              replace: `note[^1]${ADDED}`,
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([]);
      if (mode === "tracked-changes") reviewer.acceptAll();
      const saved = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        { session },
      );
      const run = addedRun(saved);
      expect(run?.formatting?.italic).toBe(true);
      expect(run?.formatting?.styleId).toBeUndefined();
      expect(run?.formatting?.vertAlign).toBeUndefined();
    },
  );
});
