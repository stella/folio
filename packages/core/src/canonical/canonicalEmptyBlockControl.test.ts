/**
 * A content control the source leaves empty (`<w:sdt/>`) stays an opaque
 * block on the canonical session: it opens, reads as on the legacy engine,
 * and a save after an edit elsewhere writes it back.
 */
import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { unzipDocx } from "../docx/unzip";

const FIXTURE = path.join(
  import.meta.dir,
  "../docx/__tests__/__fixtures__/corpus/sdt-self-closing.docx",
);

const source = (): ArrayBuffer => {
  const bytes = readFileSync(FIXTURE);
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  );
};

const controlsIn = async (buffer: ArrayBuffer): Promise<number> =>
  ((await unzipDocx(buffer)).documentXml ?? "").match(/<w:sdt[\s/>]/gu)
    ?.length ?? 0;

test("an empty content control reads alike on both engines", async () => {
  const [legacy, canonical] = await Promise.all(
    (["legacy", "canonical"] as const).map((session) =>
      FolioDocxReviewer.fromBuffer(source(), { session }),
    ),
  );
  const texts = (reviewer: FolioDocxReviewer | undefined) =>
    reviewer?.snapshot().blocks.map(({ kind, text }) => ({ kind, text }));
  expect(texts(canonical)).toEqual(texts(legacy));
});

test.each(["legacy", "canonical"] as const)(
  "an edit beside an empty content control keeps the control (%s)",
  async (session) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(source(), {
      session,
    });
    const block = reviewer
      .snapshot()
      .blocks.find(({ text }) => text === "After empty sdt.");
    if (block === undefined) throw new Error("Missing the paragraph");
    const result = reviewer.applyOperations(
      [
        {
          id: "edit",
          type: "replaceInBlock",
          blockId: block.id,
          find: "After",
          replace: "Beside",
        },
      ],
      { undefinedReferences: "refuse", mode: "direct" },
    );
    expect(result.skipped).toEqual([]);
    const saved = await reviewer.toBuffer();
    expect(await controlsIn(saved)).toBe(await controlsIn(source()));
    const reopened = await FolioDocxReviewer.fromBuffer(saved, { session });
    expect(reopened.getContentAsText()).toContain("Beside empty sdt.");
  },
);

const SESSIONS = ["legacy", "canonical"] as const;
const INSERTED = "Inserted beside the control.";

const fillerOf = (reviewer: FolioDocxReviewer) => {
  const filler = reviewer.snapshot().blocks.at(0);
  if (filler === undefined || filler.text !== "")
    throw new Error("Missing the control's filler paragraph");
  return filler;
};

const texts = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ text }) => text);

test.each(SESSIONS)(
  "a paragraph inserted after an empty content control's filler follows the control (%s)",
  async (session) => {
    for (const mode of ["direct", "tracked-changes"] as const) {
      const reviewer = await FolioDocxReviewer.fromBuffer(source(), {
        session,
      });
      const result = reviewer.applyOperations(
        [
          {
            id: "insert",
            type: "insertAfterBlock",
            blockId: fillerOf(reviewer).id,
            text: INSERTED,
          },
        ],
        { undefinedReferences: "refuse", mode },
      );
      expect(result.skipped).toEqual([]);
      expect(texts(reviewer)).toEqual(["", INSERTED, "After empty sdt."]);
      const saved = await reviewer.toBuffer();
      expect(await controlsIn(saved)).toBe(await controlsIn(source()));
      const reopened = await FolioDocxReviewer.fromBuffer(saved, { session });
      expect(texts(reopened)).toEqual(["", INSERTED, "After empty sdt."]);
      // The control stays as the source wrote it: nothing goes inside it.
      if (session === "canonical")
        expect((await unzipDocx(saved)).documentXml).toContain("<w:sdt/>");
    }
  },
);

test("a table inserted after an empty content control's filler follows the control (canonical)", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source(), {
    session: "canonical",
  });
  const result = reviewer.applyOperations(
    [
      {
        id: "table",
        type: "insertTable",
        blockId: fillerOf(reviewer).id,
        rows: [["Cell"]],
      },
    ],
    { undefinedReferences: "refuse", mode: "direct" },
  );
  expect(result.skipped).toEqual([]);
  expect(texts(reviewer)).toEqual(["", "Cell", "After empty sdt."]);
  const saved = await reviewer.toBuffer();
  expect((await unzipDocx(saved)).documentXml).toMatch(
    /<w:sdt\/>\s*<w:tbl[\s>]/u,
  );
});

test("a paragraph before an empty content control that starts the body has no paragraph to follow (canonical)", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source(), {
    session: "canonical",
  });
  const { skipped } = reviewer.applyOperations(
    [
      {
        id: "insert",
        type: "insertBeforeBlock",
        blockId: fillerOf(reviewer).id,
        text: INSERTED,
      },
    ],
    { undefinedReferences: "refuse", mode: "direct" },
  );
  expect(skipped).toMatchObject([
    {
      id: "insert",
      reason: "unsupportedBlock",
      message: expect.stringContaining("empty content control"),
    },
  ]);
});

test("an edit of an empty content control's filler names why it is refused (canonical)", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source(), {
    session: "canonical",
  });
  const blockId = fillerOf(reviewer).id;
  const { skipped } = reviewer.applyOperations(
    [
      { id: "replace", type: "replaceBlock", blockId, text: "Typed" },
      {
        id: "comment",
        type: "commentOnBlock",
        blockId,
        comment: { text: "Why?" },
      },
      { id: "delete", type: "deleteBlock", blockId },
    ],
    { undefinedReferences: "refuse", mode: "direct" },
  );
  expect(skipped.map(({ id, reason }) => ({ id, reason }))).toEqual([
    { id: "replace", reason: "unsupportedBlock" },
    { id: "comment", reason: "unsupportedBlock" },
    { id: "delete", reason: "unsupportedBlock" },
  ]);
  for (const { message } of skipped)
    expect(message).toContain("empty content control");
  expect(texts(reviewer)).toEqual(["", "After empty sdt."]);
});
