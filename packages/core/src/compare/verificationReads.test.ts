import { expect, spyOn, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { fromMarkdown } from "../markdown";
import { createDocx, ensureParaIds } from "../server";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Footnote } from "../types/document";
import { compareDocx } from "./compare";

const SESSIONS: readonly HeadlessReviewerSession[] = ["canonical"];
const NOTE_IDS = [1, 2, 3, 4, 5, 6] as const;

/** A body referencing six footnotes, each note ending in `tail`. */
const buildDocx = async (tail: string): Promise<ArrayBuffer> => {
  const model = fromMarkdown(
    NOTE_IDS.map((id) => `Paragraph ${String(id)}.`).join("\n\n"),
  );
  for (const [index, id] of NOTE_IDS.entries()) {
    const paragraph = model.package.document.content[index];
    if (paragraph?.type !== "paragraph")
      throw new Error("fixture paragraph is missing");
    paragraph.content = [
      {
        type: "run",
        formatting: {},
        content: [
          { type: "text", text: `Paragraph ${String(id)}` },
          { type: "footnoteRef", id },
          { type: "text", text: "." },
        ],
      },
    ];
  }
  model.package.footnotes = NOTE_IDS.map((id): Footnote => ({
    type: "footnote",
    id,
    noteType: "normal",
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "run",
            formatting: {},
            content: [{ type: "text", text: `Footnote ${String(id)}${tail}` }],
          },
        ],
      },
    ],
  }));
  const bytes = (await ensureParaIds(new Uint8Array(await createDocx(model))))
    .docx;
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
};

// A resolved view of a canonical session is a second session with the whole
// document resolved, so a copy for each story costs the square of their
// number. The stories of one document share a copy of each view.
test("a comparison resolves its redline once per view, however many stories it edits", async () => {
  const base = await buildDocx(".");
  const target = await buildDocx(" as amended.");
  const accept = spyOn(FolioDocxReviewer.prototype, "acceptAll");
  const reject = spyOn(FolioDocxReviewer.prototype, "rejectAll");
  try {
    const compared = await compareDocx(base, target, {
      author: "Compare",
      timestamp: "2026-11-22T00:00:00.000Z",
      session: "canonical",
    });
    if (compared.isErr()) throw compared.error;
    expect(compared.value.verification).toEqual({ status: "verified" });
    const edited = compared.value.changes.filter(
      ({ location }) => location.story.type === "footnote",
    );
    expect(edited).toHaveLength(NOTE_IDS.length);
    expect({
      accept: accept.mock.calls.length,
      reject: reject.mock.calls.length,
    }).toEqual({ accept: 1, reject: 1 });
  } finally {
    accept.mockRestore();
    reject.mockRestore();
  }
});

// A reader sharing a copy reads the document as it stands, not as it stood
// when the copy was taken.
test.each(SESSIONS)(
  "%s: a resolved story follows the edits made since it was last read",
  async (session) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocx("."), {
      session,
    });
    const story = { type: "main" } as const;
    const first = () => ({
      accepted: reviewer.readReviewedStory({ story, view: "final" })?.snapshot
        .blocks[0]?.text,
      rejected: reviewer.readReviewedStory({ story, view: "original" })
        ?.snapshot.blocks[0]?.text,
    });
    const before = first();
    expect(before.accepted).toBe(before.rejected);
    for (const text of ["Inserted once", "Inserted twice"]) {
      const blockId = reviewer.snapshotStory(story)?.blocks[0]?.id;
      if (blockId === undefined) throw new Error("fixture block is missing");
      const { skipped } = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [{ id: text, type: "insertBeforeBlock", blockId, text }],
        },
        {},
      );
      expect(skipped).toEqual([]);
      expect(first()).toEqual({ accepted: text, rejected: before.rejected });
    }
  },
);
