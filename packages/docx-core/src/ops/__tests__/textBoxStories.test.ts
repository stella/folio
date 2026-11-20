/** Paragraphs inside a text box take every operation a story's paragraphs take. */
import { describe, expect, test } from "bun:test";

import type { Document, Paragraph, ShapeContent } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation, normalizeForOps } from "../contract";
import {
  paragraphTextBoxes,
  storyBody,
  textBoxStoriesIn,
  type TextBoxStory,
} from "../stories";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  INHERIT_RUN_PROPS,
  OP_STORIES,
} from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const textBox = (content: Paragraph[]): ShapeContent => ({
  type: "shape",
  shape: {
    type: "shape",
    shapeType: "textBox",
    id: "1",
    size: { width: 2_000_000, height: 800_000 },
    textBody: { content },
  },
});

/** A body paragraph drawing one text box, between two plain paragraphs. */
const boxedDocument = (): Document =>
  normalizeForOps({
    package: {
      document: {
        content: [
          paragraph("00000001", "Before the box."),
          {
            type: "paragraph",
            paraId: "00000002",
            content: [
              {
                type: "run",
                content: [
                  { type: "text", text: "Host " },
                  textBox([
                    paragraph("00000010", "Boxed note."),
                    paragraph("00000011", "Second boxed line."),
                  ]),
                ],
              },
            ],
          },
          paragraph("00000003", "After the box."),
        ],
      },
    },
  });

const box: TextBoxStory = {
  kind: "textBox",
  host: OP_STORIES.MAIN,
  blockId: "00000002",
  index: 0,
};

const boxText = (document: Document): string[] =>
  storyBody(document, box).content.map((block) =>
    block.type === "paragraph"
      ? block.content
          .flatMap((item) => (item.type === "run" ? item.content : []))
          .map((item) => (item.type === "text" ? item.text : ""))
          .join("")
      : block.type,
  );

/** Apply, check the contract, and undo exactly through the wire. */
const exact = (document: Document, op: DocumentOp) => {
  const applied = applyDocumentOp(document, op).unwrap();
  expect(contractViolation(applied.document)).toBeUndefined();
  const inverse: DocumentOp[] = JSON.parse(JSON.stringify(applied.inverse));
  // What an operation hands back names the package story, never the box.
  expect(JSON.stringify(inverse)).not.toContain('"kind":"textBox"');
  expect(applyDocumentOps(applied.document, inverse).unwrap().document).toEqual(
    document,
  );
  return applied;
};

const at = (blockId: string, offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});

describe("text box stories", () => {
  test("a story lists the text boxes drawn in it", () => {
    const document = boxedDocument();
    expect(textBoxStoriesIn(document, OP_STORIES.MAIN)).toEqual([box]);
    const host = document.package.document.content[1];
    if (host?.type !== "paragraph") throw new Error("The host is a paragraph.");
    expect(paragraphTextBoxes(host)).toHaveLength(1);
  });

  test("text edits named through the host story land in the box and nowhere else", () => {
    const document = boxedDocument();
    const inserted = exact(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000010", 6),
      text: "short ",
      runProps: INHERIT_RUN_PROPS,
    });
    expect(boxText(inserted.document)).toEqual([
      "Boxed short note.",
      "Second boxed line.",
    ]);
    const [before, , after] = inserted.document.package.document.content;
    expect(before).toBe(document.package.document.content[0]);
    expect(after).toBe(document.package.document.content[2]);
    const deleted = exact(document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at("00000010", 0),
      to: at("00000010", 6),
    });
    expect(boxText(deleted.document)).toEqual(["note.", "Second boxed line."]);
  });

  test("block operations keep the box's own paragraph list", () => {
    const document = boxedDocument();
    const inserted = exact(document, {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "00000011" },
      blocks: [paragraph("00000012", "Third boxed line.")],
    });
    expect(boxText(inserted.document)).toEqual([
      "Boxed note.",
      "Second boxed line.",
      "Third boxed line.",
    ]);
    const split = exact(document, {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: at("00000010", 6),
      newBlockId: "00000013",
      newHalf: "second",
    });
    expect(boxText(split.document)).toEqual([
      "Boxed ",
      "note.",
      "Second boxed line.",
    ]);
    const joined = exact(document, {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      nextBlockId: "00000011",
    });
    expect(boxText(joined.document)).toEqual(["Boxed note.Second boxed line."]);
  });

  test("tracked edits record revisions in the box, which resolve through the host story", () => {
    const document = boxedDocument();
    const tracked = exact(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000010", 6),
      text: "short ",
      runProps: INHERIT_RUN_PROPS,
      revision: { id: 7, author: "Reviewer", date: "2026-11-20T00:00:00Z" },
    });
    expect(tracked.revisions).toEqual([7]);
    for (const decision of ["accept", "reject"] as const) {
      const resolved = exact(tracked.document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [7],
        decision,
      });
      expect(boxText(resolved.document)).toEqual([
        decision === "accept" ? "Boxed short note." : "Boxed note.",
        "Second boxed line.",
      ]);
      expect(JSON.stringify(resolved.document)).not.toContain('"id":7');
    }
  });

  test("a revision resolution spanning the body and a box resolves both", () => {
    const document = boxedDocument();
    const revision = (id: number) => ({
      id,
      author: "Reviewer",
      date: "2026-11-20T00:00:00Z",
    });
    const tracked = applyDocumentOps(document, [
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: at("00000010", 0),
        text: "New ",
        runProps: INHERIT_RUN_PROPS,
        revision: revision(7),
      },
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: at("00000001", 0),
        text: "New ",
        runProps: INHERIT_RUN_PROPS,
        revision: revision(8),
      },
    ]).unwrap();
    const accepted = exact(tracked.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [7, 8],
      decision: "accept",
    });
    expect(boxText(accepted.document)[0]).toBe("New Boxed note.");
    expect(JSON.stringify(accepted.document)).not.toContain('"author"');
  });

  test("an operation naming paragraphs of the body and of a box finds no paragraph", () => {
    const document = boxedDocument();
    const result = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at("00000001", 0),
      to: at("00000010", 2),
    });
    expect(result.isErr()).toBe(true);
  });

  test("comments open, resolve and go on a text box's words", () => {
    const document = boxedDocument();
    const commented = exact(document, {
      type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
      comment: {
        id: 100,
        author: "Author",
        content: [paragraph("00000020", "root")],
      },
      anchor: {
        kind: "range",
        from: at("00000010", 0),
        to: at("00000010", 5),
      },
    });
    expect(JSON.stringify(storyBody(commented.document, box))).toContain(
      '"commentRangeStart","id":100',
    );
    const resolved = exact(commented.document, {
      type: DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION,
      id: 100,
      status: "resolved",
    });
    const deleted = exact(resolved.document, {
      type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
      id: 100,
      scope: "thread",
    });
    expect(storyBody(deleted.document, box)).toEqual(storyBody(document, box));
  });

  test("a tracked paragraph after a box's last paragraph reports every revision it records", () => {
    const document = boxedDocument();
    const op: DocumentOp = {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "00000011" },
      blocks: [paragraph("00000012", "Third boxed line.")],
      revision: { id: 7, author: "Reviewer", date: "2026-11-20T00:00:00Z" },
      newIds: { revision: [8, 9] },
    };
    const own = applyDocumentOp(
      normalizeForOps({
        package: { document: { content: storyBody(document, box).content } },
      }),
      op,
    ).unwrap();
    const inBox = exact(document, op);
    expect(inBox.revisions).toEqual(own.revisions);
    expect(inBox.revisions.length).toBeGreaterThan(1);
  });
});
