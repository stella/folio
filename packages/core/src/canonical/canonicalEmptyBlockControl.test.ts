/**
 * A content control the source leaves empty (`<w:sdt/>`) stays an opaque
 * block on the canonical session: it opens, reads as on the legacy engine,
 * and a save after an edit elsewhere writes it back.
 */
import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import type { BlockContent } from "@stll/docx-core/model";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import { unzipDocx } from "../docx/unzip";
import { fromMarkdown } from "../markdown/fromMarkdown";

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

test.each(["canonical"] as const)(
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

const SESSIONS = ["canonical"] as const;
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

const BEFORE_START = [
  {
    name: "a paragraph",
    operation: (blockId: string) =>
      ({
        id: "insert",
        type: "insertBeforeBlock",
        blockId,
        text: INSERTED,
      }) as const,
    inserted: INSERTED,
  },
  {
    name: "a table",
    operation: (blockId: string) =>
      ({
        id: "insert",
        type: "insertTable",
        blockId,
        position: "before",
        rows: [["Cell"]],
      }) as const,
    inserted: "Cell",
  },
] as const;

for (const session of SESSIONS)
  for (const mode of ["direct", "tracked-changes"] as const)
    for (const { name, operation, inserted } of BEFORE_START)
      test(`${name} inserted before an empty content control that starts the body goes before it (${session}, ${mode})`, async () => {
        const reviewer = await FolioDocxReviewer.fromBuffer(source(), {
          session,
        });
        const result = reviewer.applyOperations(
          [operation(fillerOf(reviewer).id)],
          { undefinedReferences: "refuse", mode },
        );
        expect(result.skipped).toEqual([]);
        const expected = [inserted, "", "After empty sdt."];
        expect(texts(reviewer)).toEqual(expected);
        const saved = await reviewer.toBuffer();
        expect(await controlsIn(saved)).toBe(await controlsIn(source()));
        const reopened = await FolioDocxReviewer.fromBuffer(saved, {
          session,
        });
        expect(texts(reopened)).toEqual(expected);
        if (session === "canonical")
          expect((await unzipDocx(saved)).documentXml).toMatch(
            /<w:body>\s*<w:(p|tbl)[\s>](?:.|\s)*?<w:sdt\/>/u,
          );
        if (mode === "direct") return;
        const rejected = await FolioDocxReviewer.fromBuffer(saved, {
          session,
        });
        rejected.rejectAll();
        expect(texts(rejected)).toEqual(["", "After empty sdt."]);
        reopened.acceptAll();
        expect(texts(reopened)).toEqual(expected);
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

for (const session of SESSIONS)
  for (const mode of ["direct", "tracked-changes"] as const)
    test(`blocks a batch inserts before an empty content control that starts the body keep the batch's order (${session}, ${mode})`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(source(), {
        session,
      });
      const blockId = fillerOf(reviewer).id;
      const result = reviewer.applyOperations(
        [
          { id: "first", type: "insertBeforeBlock", blockId, text: "First." },
          {
            id: "table",
            type: "insertTable",
            blockId,
            position: "before",
            rows: [["Cell"]],
          },
          { id: "last", type: "insertBeforeBlock", blockId, text: "Last." },
        ],
        { undefinedReferences: "refuse", mode },
      );
      expect(result.skipped).toEqual([]);
      const expected = ["First.", "Cell", "Last.", "", "After empty sdt."];
      expect(texts(reviewer)).toEqual(expected);
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
        { session },
      );
      expect(texts(reopened)).toEqual(expected);
    });

/** A document from `markdown` with an empty content control put in by `place`. */
const withEmptyControl = async (
  markdown: string,
  place: (content: BlockContent[], control: BlockContent) => void,
) => {
  const document = fromMarkdown(markdown);
  place(document.package.document.content, {
    type: "blockSdt",
    properties: { sdtType: "richText" },
    content: [],
  });
  return await createDocx(document);
};

const PLACES = [
  {
    name: "first in a table cell goes beside the table",
    input: () =>
      withEmptyControl("| H |\n|---|\n| A |\n\nAfter.", (content, control) => {
        const table = content.find((block) => block.type === "table");
        table?.rows.at(0)?.cells.at(0)?.content.unshift(control);
      }),
    expected: ["Heading", "", "H", "A", "After."],
    // Tracked, rejecting the paragraph would run its mark into a cell that
    // starts with a control, which docx-core refuses for any table.
    modes: ["direct"],
  },
  {
    name: "after a table goes between them",
    input: () =>
      withEmptyControl("| H |\n|---|\n| A |\n\nAfter.", (content, control) => {
        content.splice(1, 0, control);
      }),
    expected: ["H", "A", "Heading", "", "After."],
    modes: ["direct", "tracked-changes"],
  },
] as const;

for (const session of SESSIONS)
  for (const { name, input, expected, modes } of PLACES)
    for (const mode of modes)
      test(`a paragraph inserted before an empty content control ${name} (${session}, ${mode})`, async () => {
        const reviewer = await FolioDocxReviewer.fromBuffer(await input(), {
          session,
        });
        const filler = reviewer
          .snapshot()
          .blocks.find(({ text }) => text === "");
        if (filler === undefined) throw new Error("Missing the filler");
        const result = reviewer.applyOperations(
          [
            {
              id: "insert",
              type: "insertBeforeBlock",
              blockId: filler.id,
              text: "Heading",
            },
          ],
          { undefinedReferences: "refuse", mode },
        );
        expect(result.skipped).toEqual([]);
        expect(texts(reviewer)).toEqual([...expected]);
        const reopened = await FolioDocxReviewer.fromBuffer(
          await reviewer.toBuffer(),
          { session },
        );
        expect(texts(reopened)).toEqual([...expected]);
      });
