/**
 * What replaces a range takes no comment whose end the range reaches: the new
 * words go past that end, so a replacement never stretches a comment over
 * words it did not cover. A comment starting in the range keeps the new words,
 * as accepting the tracked replacement leaves them inside it. Tracked, the
 * struck words keep their ranges; direct, a range the replacement covers
 * wholly goes with its words.
 */
import { describe, expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { allocateEditorIntentIds, compileEditorIntent } from "../editorIntent";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const PARA_ID = "12345678";
const DATE = "2026-01-01T00:00:00Z";

const run = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

/** "The goods on time and in order." with a comment over "on time and". */
const commented = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: PARA_ID,
          content: [
            run("The goods "),
            { type: "commentRangeStart", id: 1 },
            run("on time and"),
            { type: "commentRangeEnd", id: 1 },
            run(" in order."),
            { type: "run", content: [{ type: "commentReference", id: 1 }] },
          ],
        },
      ],
      comments: [
        {
          id: 1,
          author: "A",
          date: DATE,
          anchorKind: "range",
          content: [
            { type: "paragraph", paraId: "0A000001", content: [run("Note.")] },
          ],
        },
      ],
    },
  },
});

/** The paragraph as text, `[`/`]` for the comment's range, `-…-`/`+…+` for revisions. */
const shape = (document: Document): string => {
  const [block] = document.package.document.content;
  if (block?.type !== "paragraph") return "";
  const write = (nodes: readonly ParagraphContent[]): string =>
    nodes
      .map((node) => {
        switch (node.type) {
          case "commentRangeStart":
            return "[";
          case "commentRangeEnd":
            return "]";
          case "run":
            return node.content
              .map((leaf) => (leaf.type === "text" ? leaf.text : ""))
              .join("");
          case "deletion":
            return `-${write(node.content as ParagraphContent[])}-`;
          case "insertion":
            return `+${write(node.content as ParagraphContent[])}+`;
          default:
            return "";
        }
      })
      .join("");
  return write(block.content);
};

const replace = (
  from: number,
  to: number,
  text: string,
  tracked: boolean,
): Document => {
  const document = commented();
  const at = (offset: number) => ({
    story: OP_STORIES.MAIN,
    blockId: PARA_ID,
    offset,
  });
  const intent = {
    type: "replaceText" as const,
    from: at(from),
    to: at(to),
    text,
  };
  const ids = allocateEditorIntentIds(document, intent);
  const compiled = compileEditorIntent(document, {
    intent,
    mode: tracked
      ? {
          type: "suggesting",
          revision: { id: ids.revisionId, author: "B", date: DATE },
          newIds: ids.newIds,
        }
      : { type: "editing", newIds: ids.newIds },
  }).unwrap();
  return applyDocumentOps(document, compiled.ops).unwrap().document;
};

// Offsets in "The goods on time and in order.": the comment covers 10..21.
const CASES = [
  { name: "the commented words", from: 10, to: 21, text: "at once" },
  { name: "more than them", from: 4, to: 24, text: "items promptly" },
  { name: "words inside them", from: 13, to: 17, text: "moment" },
  { name: "across their start", from: 4, to: 12, text: "items at" },
  { name: "across their end", from: 18, to: 24, text: "plus at" },
  { name: "their start onwards", from: 10, to: 17, text: "at once" },
  { name: "up to their end", from: 13, to: 21, text: "moment plus" },
] as const;

const TRACKED: Record<(typeof CASES)[number]["name"], string> = {
  "the commented words": "The goods [-on time and-]+at once+ in order.",
  "more than them": "The -goods -[-on time and-]- in-+items promptly+ order.",
  "words inside them": "The goods [on -time-+moment+ and] in order.",
  "across their start": "The -goods -[-on-+items at+ time and] in order.",
  "across their end": "The goods [on time -and-]- in-+plus at+ order.",
  "their start onwards": "The goods [-on time-+at once+ and] in order.",
  "up to their end": "The goods [on -time and-]+moment plus+ in order.",
};

const DIRECT: Record<(typeof CASES)[number]["name"], string> = {
  "the commented words": "The goods at once in order.",
  "more than them": "The items promptly order.",
  "words inside them": "The goods [on moment and] in order.",
  "across their start": "The [items at time and] in order.",
  "across their end": "The goods [on time ]plus at order.",
  "their start onwards": "The goods [at once and] in order.",
  "up to their end": "The goods [on ]moment plus in order.",
};

describe("a replacement beside a comment's range", () => {
  for (const { name, from, to, text } of CASES) {
    test(`tracked over ${name}`, () => {
      expect(shape(replace(from, to, text, true))).toBe(TRACKED[name]);
    });
    test(`direct over ${name}`, () => {
      const replaced = replace(from, to, text, false);
      expect(shape(replaced)).toBe(DIRECT[name]);
      expect((replaced.package.document.comments ?? []).length).toBe(
        DIRECT[name].includes("[") ? 1 : 0,
      );
    });
  }
});

test("comment ends at a content control's end all close it: the replacement stays in one control", () => {
  const control = (content: ParagraphContent[]): ParagraphContent => ({
    type: "inlineSdt",
    properties: { sdtType: "richText", id: 5 },
    content,
  });
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: PARA_ID,
            content: [
              { type: "commentRangeStart", id: 1 },
              { type: "commentRangeStart", id: 2 },
              control([
                run("old"),
                { type: "commentRangeEnd", id: 1 },
                { type: "commentRangeEnd", id: 2 },
              ]),
            ],
          },
        ],
        comments: [1, 2].map((id) => ({
          id,
          author: "A",
          date: DATE,
          anchorKind: "range" as const,
          content: [
            {
              type: "paragraph" as const,
              paraId: `0A00000${id}`,
              content: [run("Note.")],
            },
          ],
        })),
      },
    },
  };
  const at = (offset: number) => ({
    story: OP_STORIES.MAIN,
    blockId: PARA_ID,
    offset,
  });
  const intent = {
    type: "replaceText" as const,
    from: at(0),
    to: at(3),
    text: "new",
  };
  const ids = allocateEditorIntentIds(document, intent);
  const outcome = (tracked: boolean) => {
    const compiled = compileEditorIntent(document, {
      intent,
      mode: tracked
        ? {
            type: "suggesting",
            revision: { id: ids.revisionId, author: "B", date: DATE },
            newIds: ids.newIds,
          }
        : { type: "editing", newIds: ids.newIds },
    }).unwrap();
    const edited = applyDocumentOps(document, compiled.ops).unwrap().document;
    if (!tracked) return edited;
    const [block] = edited.package.document.content;
    const [wrapper] =
      block?.type === "paragraph" ? block.content.slice(-1) : [];
    const revisionIds = (
      wrapper?.type === "inlineSdt" ? wrapper.content : []
    ).flatMap((node) =>
      node.type === "insertion" || node.type === "deletion"
        ? [node.info.id]
        : [],
    );
    expect(revisionIds.length).toBe(2);
    return applyDocumentOps(edited, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds,
        decision: REVISION_DECISIONS.ACCEPT,
      },
    ]).unwrap().document;
  };
  const expected = [
    { type: "commentRangeStart", id: 1 },
    { type: "commentRangeStart", id: 2 },
    control([
      run("new"),
      { type: "commentRangeEnd", id: 1 },
      { type: "commentRangeEnd", id: 2 },
    ]),
  ];
  for (const tracked of [false, true]) {
    const [block] = outcome(tracked).package.document.content;
    expect(block?.type === "paragraph" ? block.content : []).toEqual(expected);
  }
});
