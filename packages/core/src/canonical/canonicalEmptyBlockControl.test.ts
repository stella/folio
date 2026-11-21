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
