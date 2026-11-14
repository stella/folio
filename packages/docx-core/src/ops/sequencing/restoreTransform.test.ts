import { describe, expect, mock, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  Document,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
} from "../../model/document";
import { equalForStaleness, structurallyEqual } from "../equality";
import { compareGaps, partitionContent } from "../leaves";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type OpStory,
  type TextPosition,
} from "../types";
import { applyBatch } from "./applyBatch";
import type { DocumentBatch, SequencedBatch } from "./envelope";
import * as restoreTransformModule from "./restoreTransform";
import { restoredSpan } from "./restoreTransform";
import { transformBatch } from "./transform";

const BLOCK = "00000001";

const run = (content: RunContent[], bold = false): Run =>
  bold
    ? { type: "run", formatting: { bold: true }, content }
    : { type: "run", content };
const text = (value: string): RunContent => ({ type: "text", text: value });

// ---------------------------------------------------------------------------
// The span
// ---------------------------------------------------------------------------

const leafArbitrary = fc.constantFrom<RunContent>(
  text("ab"),
  text("c"),
  text("de"),
  { type: "tab" },
  { type: "renderedPageBreak" },
);

const runArbitrary = fc
  .tuple(fc.array(leafArbitrary, { minLength: 1, maxLength: 3 }), fc.boolean())
  .map(([content, bold]) => run(content, bold));

const contentArbitrary = fc.array(runArbitrary, { minLength: 1, maxLength: 4 });

/** One local change: new text, a run's formatting, a run cut or merged, a run added or removed. */
const changed = (
  content: readonly Run[],
  kind: number,
  index: number,
  replacement: Run,
): Run[] => {
  const at = index % content.length;
  const target = content[at] ?? replacement;
  const out = [...content];
  switch (kind % 6) {
    case 0:
      out[at] = {
        ...target,
        content: [text("xy"), ...target.content.slice(1)],
      };
      return out;
    case 1:
      out[at] =
        target.formatting === undefined
          ? { ...target, formatting: { bold: true } }
          : { type: "run", content: target.content };
      return out;
    case 2: {
      if (target.content.length < 2) return out;
      out.splice(
        at,
        1,
        { ...target, content: target.content.slice(0, 1) },
        { ...target, content: target.content.slice(1) },
      );
      return out;
    }
    case 3: {
      const next = content[at + 1];
      if (
        next === undefined ||
        !structurallyEqual(next.formatting, target.formatting)
      )
        return out;
      out.splice(at, 2, {
        ...target,
        content: [...target.content, ...next.content],
      });
      return out;
    }
    case 4:
      out.splice(at, 0, replacement);
      return out;
    default:
      if (out.length > 1) out.splice(at, 1);
      return out;
  }
};

/** The content outside the span: before it, and after it. */
const outside = (
  content: readonly ParagraphContent[],
  from: Parameters<typeof partitionContent>[1][number],
  to: Parameters<typeof partitionContent>[1][number],
) => {
  const [before = [], , after = []] = partitionContent(content, [from, to]);
  return { before, after };
};

describe("the span a restoration changes", () => {
  test(
    "is the stretch outside of which both contents are the same, and no smaller",
    () => {
      fc.assert(
        fc.property(
          contentArbitrary,
          fc.nat(),
          fc.nat(),
          runArbitrary,
          (expected, kind, index, replacement) => {
            const content = changed(expected, kind, index, replacement);
            const span = restoredSpan(expected, content);
            // Sound: everything outside the span is the same.
            expect(outside(expected, span.from, span.to)).toEqual(
              outside(content, span.from, span.restoredTo),
            );
            // Minimal: moving either end inwards takes a difference outside.
            const startShrinks =
              span.from.offset < span.to.offset &&
              span.from.offset < span.restoredTo.offset;
            if (startShrinks) {
              const from = { offset: span.from.offset + 1, zeroWidthBefore: 0 };
              expect(
                structurallyEqual(
                  outside(expected, from, span.to),
                  outside(content, from, span.restoredTo),
                ),
              ).toBe(false);
            }
            const endShrinks =
              span.to.offset > span.from.offset &&
              span.restoredTo.offset > span.from.offset;
            const to = { offset: span.to.offset - 1, zeroWidthBefore: 0 };
            const restoredTo = {
              offset: span.restoredTo.offset - 1,
              zeroWidthBefore: 0,
            };
            if (
              endShrinks &&
              compareGaps(span.from, to) <= 0 &&
              compareGaps(span.from, restoredTo) <= 0
            ) {
              expect(
                structurallyEqual(
                  outside(expected, span.from, to),
                  outside(content, span.from, restoredTo),
                ),
              ).toBe(false);
            }
          },
        ),
        propertyConfig({ numRuns: 300 }),
      );
    },
    propertyTestTimeout(60_000),
  );

  test("of a change in a run's own fields alone is that run's leaves", () => {
    const expected = [
      run([text("ab")]),
      run([text("cd")], true),
      run([text("ef")]),
    ];
    const content = [run([text("ab")]), run([text("cd")]), run([text("ef")])];
    expect(restoredSpan(expected, content)).toEqual({
      from: { offset: 2, zeroWidthBefore: 0 },
      to: { offset: 4, zeroWidthBefore: 0 },
      restoredTo: { offset: 4, zeroWidthBefore: 0 },
    });
  });

  test("of a run cut in two alone is the empty stretch at the cut", () => {
    const expected = [run([text("abcd")])];
    const content = [run([text("ab")]), run([text("cd")])];
    expect(restoredSpan(expected, content)).toEqual({
      from: { offset: 2, zeroWidthBefore: 0 },
      to: { offset: 2, zeroWidthBefore: 0 },
      restoredTo: { offset: 2, zeroWidthBefore: 0 },
    });
  });
});

// ---------------------------------------------------------------------------
// Restorations against edits of their paragraph
// ---------------------------------------------------------------------------

type Place = {
  story: OpStory;
  document: (content: ParagraphContent[]) => Document;
};

const paragraphOf = (content: ParagraphContent[]): Paragraph => ({
  type: "paragraph",
  paraId: BLOCK,
  content,
});

const PLACES: Record<"main" | "tableCell" | "footnote", Place> = {
  main: {
    story: OP_STORIES.MAIN,
    document: (content) => ({
      package: { document: { content: [paragraphOf(content)] } },
    }),
  },
  tableCell: {
    story: OP_STORIES.MAIN,
    document: (content) => ({
      package: {
        document: {
          content: [
            {
              type: "table",
              rows: [
                {
                  type: "tableRow",
                  cells: [
                    { type: "tableCell", content: [paragraphOf(content)] },
                  ],
                },
              ],
            },
          ],
        },
      },
    }),
  },
  footnote: {
    story: { kind: "footnote", id: 1 },
    document: (content) => ({
      package: {
        document: {
          content: [{ type: "paragraph", paraId: "0000000F", content: [] }],
        },
        footnotes: [
          { type: "footnote", id: 1, content: [paragraphOf(content)] },
        ],
      },
    }),
  },
};

// "Hello " plain, "world" bold, " again" plain: the restoration unbolds "world".
const before: ParagraphContent[] = [
  run([text("Hello ")]),
  run([text("world")], true),
  run([text(" again")]),
];
const after: ParagraphContent[] = [run([text("Hello world again")])];

const batchOf = (opId: string, op: DocumentOp): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor: opId,
  baseRev: 0,
  ops: [op],
});

/** `first` sequenced, then `second` rebased over it and applied. */
const inOrder = (
  base: Document,
  first: DocumentBatch,
  second: DocumentBatch,
): Document | string => {
  const applied = applyBatch(base, first.ops);
  if (applied.isErr()) return applied.error.message;
  const sequenced: SequencedBatch = {
    ...first,
    revision: 1,
    effects: applied.value.effects,
  };
  const rebased = transformBatch(second, [sequenced]);
  if (rebased.isErr()) return rebased.error.message;
  const next = applyBatch(applied.value.document, rebased.value.ops);
  return next.isErr() ? next.error.message : next.value.document;
};

const restoration = (story: OpStory): DocumentBatch =>
  batchOf("restoration", {
    type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
    story,
    blockId: BLOCK,
    expected: before,
    content: after,
  });

const position = (story: OpStory, offset: number): TextPosition => ({
  story,
  blockId: BLOCK,
  offset,
});

const EDITS = {
  "typing before the span": (story) => ({
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: position(story, 2),
    text: "XY",
    runProps: "inherit",
  }),
  "typing after the span": (story) => ({
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: position(story, 14),
    text: "XY",
    runProps: "inherit",
  }),
  "a deletion after the span": (story) => ({
    type: DOCUMENT_OP_TYPES.DELETE_RANGE,
    from: position(story, 12),
    to: position(story, 15),
  }),
  "a format before the span": (story) => ({
    type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
    from: position(story, 0),
    to: position(story, 3),
    patch: { italic: true },
  }),
} satisfies Record<string, (story: OpStory) => DocumentOp>;

describe("a paragraph restoration and an edit of its paragraph", () => {
  for (const [placeName, place] of Object.entries(PLACES))
    for (const [editName, edit] of Object.entries(EDITS))
      test(`${editName} passes it in either order (${placeName})`, () => {
        const base = place.document(before);
        const restore = restoration(place.story);
        const change = batchOf("edit", edit(place.story));
        const restoredFirst = inOrder(base, restore, change);
        const editedFirst = inOrder(base, change, restore);
        expect(typeof restoredFirst).not.toBe("string");
        expect(typeof editedFirst).not.toBe("string");
        if (
          typeof restoredFirst === "string" ||
          typeof editedFirst === "string"
        )
          return;
        expect(equalForStaleness(restoredFirst, editedFirst)).toBe(true);
      });

  test("typing inside the span conflicts with it in either order", () => {
    const base = PLACES.main.document(before);
    const restore = restoration(OP_STORIES.MAIN);
    const typing = batchOf("edit", {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: position(OP_STORIES.MAIN, 8),
      text: "XY",
      runProps: "inherit",
    });
    const message =
      "A concurrent edit reached the content a restoration changes.";
    expect(inOrder(base, restore, typing)).toBe(message);
    expect(inOrder(base, typing, restore)).toBe(message);
  });

  test("a restore transform that passes nothing is caught", () => {
    // The test-only break: every restoration conflicts with every edit of
    // its paragraph, as a block footprint alone decides.
    const original = { ...restoreTransformModule };
    mock.module("./restoreTransform", () => ({
      ...original,
      transformRestorePair: () => undefined,
    }));
    try {
      const base = PLACES.main.document(before);
      const typing = batchOf(
        "edit",
        EDITS["typing after the span"](OP_STORIES.MAIN),
      );
      expect(inOrder(base, restoration(OP_STORIES.MAIN), typing)).toBe(
        "A concurrent operation changed a paragraph this operation addresses.",
      );
    } finally {
      mock.module("./restoreTransform", () => original);
    }
  });
});
