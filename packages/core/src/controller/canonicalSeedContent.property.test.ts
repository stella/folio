/**
 * The seed check judges each block once and reuses the judgment while the
 * block stays the same object, as every block an operation does not name
 * does. Along a chain of edits that share blocks, its verdict equals one walk
 * of every story. The check of an edit's result judges only the blocks the
 * edit replaced, and the comments only when it changed their anchors or
 * definitions; after a version the session could hold, its verdict equals
 * one walk too.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import type { BlockContent, Paragraph } from "@stll/docx-core/model";
import { OP_STORIES } from "@stll/docx-core/ops";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { table } from "../docx/server/build";
import type { Document } from "../types/document";
import {
  unsupportedEditReason,
  unsupportedSeedReason,
  unsupportedSeedReasonWholeDocument,
} from "./canonicalSeedContent";

// Synchronous generated edit chains use a 15-second base budget.
setDefaultTimeout(propertyTestTimeout(15_000));

// A tab in a text node is one of the inline contents the session refuses.
const texts = fc.constantFrom("a", "b", "c\td");

const commentIds = fc.constantFrom(1, 2, 3);

/** A comment anchor in a paragraph: a range boundary or a reference. */
const inlineAnchorOf = fc.record({
  type: fc.constantFrom(
    "commentRangeStart" as const,
    "commentRangeEnd" as const,
    "commentReference" as const,
  ),
  id: commentIds,
});

const paragraphOf = fc
  .array(fc.oneof({ weight: 3, arbitrary: texts }, inlineAnchorOf), {
    maxLength: 2,
  })
  .map((parts): Paragraph => ({
    type: "paragraph",
    content: parts.map((part) =>
      typeof part === "string"
        ? { type: "run", content: [{ type: "text", text: part }] }
        : part,
    ),
  }));

const blockOf: fc.Arbitrary<BlockContent> = fc.oneof(
  { weight: 4, arbitrary: paragraphOf },
  {
    weight: 1,
    arbitrary: fc
      .array(paragraphOf, { minLength: 1, maxLength: 2 })
      .map((content): BlockContent => {
        const cells = table({ rows: [["x"]] });
        const cell = cells.rows.at(0)?.cells.at(0);
        if (cell !== undefined) cell.content = content;
        return cells;
      }),
  },
  {
    weight: 1,
    arbitrary: fc
      .array(paragraphOf, { minLength: 1, maxLength: 2 })
      .map((content): BlockContent => ({
        type: "blockSdt",
        properties: { sdtType: "richText" },
        content,
      })),
  },
  // A comment range boundary between blocks.
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constantFrom(
        "commentRangeStart" as const,
        "commentRangeEnd" as const,
      ),
      id: commentIds,
    }),
  },
  // A block the session cannot hold.
  {
    weight: 1,
    arbitrary: fc
      .array(paragraphOf, { maxLength: 1 })
      .map((content): BlockContent => ({
        type: "blockCustomXml",
        openingXml: "<w:customXml>",
        closingXml: "</w:customXml>",
        content,
      })),
  },
);

type Edit =
  | { type: "replace"; at: number; block: BlockContent }
  | { type: "insert"; at: number; block: BlockContent }
  | { type: "delete"; at: number }
  | { type: "comments"; ids: number[] };

const editOf: fc.Arbitrary<Edit> = fc.oneof(
  fc.record({
    type: fc.constant("replace" as const),
    at: fc.nat(),
    block: blockOf,
  }),
  fc.record({
    type: fc.constant("insert" as const),
    at: fc.nat(),
    block: blockOf,
  }),
  fc.record({ type: fc.constant("delete" as const), at: fc.nat() }),
  fc.record({
    type: fc.constant("comments" as const),
    ids: fc.uniqueArray(commentIds, { maxLength: 3 }),
  }),
);

const commentsOf = (ids: readonly number[]) =>
  ids.map((id) => ({ id, author: "Reviewer", content: [] }));

const documentOf = (
  body: BlockContent[],
  note: BlockContent[],
  comments?: Document["package"]["document"]["comments"],
): Document => ({
  package: {
    document: {
      content: body,
      ...(comments === undefined ? {} : { comments }),
    },
    footnotes: [{ type: "footnote", id: 1, content: note }],
  },
});

/** A new document sharing every block the edit does not touch. */
const edited = (document: Document, edit: Edit, inNote: boolean): Document => {
  const note = document.package.footnotes?.at(0)?.content ?? [];
  const comments = document.package.document.comments;
  if (edit.type === "comments")
    return documentOf(
      document.package.document.content,
      note,
      commentsOf(edit.ids),
    );
  const blocks = [...(inNote ? note : document.package.document.content)];
  const at = blocks.length === 0 ? 0 : edit.at % blocks.length;
  if (edit.type === "replace" && blocks.length > 0)
    blocks.splice(at, 1, edit.block);
  else if (edit.type === "insert") blocks.splice(at, 0, edit.block);
  else if (edit.type === "delete") blocks.splice(at, 1);
  return inNote
    ? documentOf(document.package.document.content, blocks, comments)
    : documentOf(blocks, note, comments);
};

/** The stories whose blocks differ between two versions. */
const changedStoriesOf = (previous: Document, document: Document) => [
  ...(previous.package.document.content === document.package.document.content
    ? []
    : [OP_STORIES.MAIN]),
  ...(previous.package.footnotes?.at(0)?.content ===
  document.package.footnotes?.at(0)?.content
    ? []
    : [{ kind: "footnote" as const, id: 1 }]),
];

describe("canonical seed check", () => {
  test("judges shared blocks as one walk of every story does", () => {
    fc.assert(
      fc.property(
        fc.array(blockOf, { maxLength: 5 }),
        fc.array(blockOf, { maxLength: 2 }),
        fc.array(fc.tuple(editOf, fc.boolean()), {
          minLength: 1,
          maxLength: 8,
        }),
        (body, note, edits) => {
          let document = documentOf(body, note);
          expect(unsupportedSeedReason(document)).toBe(
            unsupportedSeedReasonWholeDocument(document),
          );
          for (const [edit, inNote] of edits) {
            document = edited(document, edit, inNote);
            expect(unsupportedSeedReason(document)).toBe(
              unsupportedSeedReasonWholeDocument(document),
            );
          }
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("judges an edit's result as one walk of every story does", () => {
    let judged = 0;
    fc.assert(
      fc.property(
        fc.array(blockOf, { maxLength: 5 }),
        fc.array(blockOf, { maxLength: 2 }),
        fc.uniqueArray(commentIds, { maxLength: 3 }),
        fc.array(fc.tuple(editOf, fc.boolean()), {
          minLength: 1,
          maxLength: 10,
        }),
        (body, note, ids, edits) => {
          const plain = (): Paragraph => ({ type: "paragraph", content: [] });
          let document = documentOf(body, note, commentsOf(ids));
          // The session holds only documents the seed check accepts.
          if (unsupportedSeedReasonWholeDocument(document) !== null)
            document = documentOf([plain()], [plain()], commentsOf(ids));
          expect(unsupportedSeedReasonWholeDocument(document)).toBeNull();
          for (const [edit, inNote] of edits) {
            const previous = document;
            const next = edited(document, edit, inNote);
            const stories = changedStoriesOf(previous, next);
            const verdict = unsupportedSeedReasonWholeDocument(next, stories);
            expect(
              unsupportedEditReason({ previous, document: next, stories }),
            ).toBe(verdict);
            judged += 1;
            // A refused edit leaves the session on the document before it.
            if (unsupportedSeedReasonWholeDocument(next) === null)
              document = next;
          }
        },
      ),
      propertyConfig({ numRuns: 600 }),
    );
    expect(judged).toBeGreaterThan(300);
  });

  describe("judges again an edit that changes the paragraph a marker between blocks attaches to", () => {
    const text = (value: string): Paragraph => ({
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: value }] }],
    });
    const start: BlockContent = { type: "commentRangeStart", id: 1 };
    const end: BlockContent = { type: "commentRangeEnd", id: 1 };
    const expectJudgedAsOneWalk = (
      previous: Document,
      document: Document,
    ): void => {
      expect(unsupportedSeedReasonWholeDocument(previous)).toBeNull();
      const stories = changedStoriesOf(previous, document);
      const verdict = unsupportedSeedReasonWholeDocument(document, stories);
      expect(verdict).not.toBeNull();
      expect(unsupportedEditReason({ previous, document, stories })).toBe(
        verdict,
      );
    };

    test("deleting the empty paragraph a range opens and closes on", () => {
      // The range starts and ends on the empty paragraph; without it, its
      // end closes on the paragraph before while its start waits for one.
      const first = text("a");
      const empty: Paragraph = { type: "paragraph", content: [] };
      const previous = documentOf(
        [first, start, empty, end],
        [{ type: "paragraph", content: [] }],
        commentsOf([1]),
      );
      expectJudgedAsOneWalk(
        previous,
        edited(previous, { type: "delete", at: 2 }, false),
      );
    });

    test("moving a range start past the empty paragraph it opened on", () => {
      // Both runs hold the same anchors, but the moved start is held past
      // the paragraph, after the end that closes on it.
      const first = text("a");
      const empty: Paragraph = { type: "paragraph", content: [] };
      const note: BlockContent[] = [{ type: "paragraph", content: [] }];
      const previous = documentOf(
        [first, start, empty, end],
        note,
        commentsOf([1]),
      );
      // The same comments: only the anchors decide that the check runs.
      expectJudgedAsOneWalk(
        previous,
        documentOf(
          [first, empty, start, end],
          note,
          previous.package.document.comments,
        ),
      );
    });
  });

  test("judges again an edit that turns an anchor into another kind with the same id", () => {
    const anchor = (
      type: "commentRangeStart" | "commentRangeEnd" | "commentReference",
    ): Paragraph => ({ type: "paragraph", content: [{ type, id: 1 }] });
    const previous = documentOf(
      [anchor("commentRangeStart"), anchor("commentRangeEnd")],
      [{ type: "paragraph", content: [] }],
      commentsOf([1]),
    );
    expect(unsupportedSeedReasonWholeDocument(previous)).toBeNull();
    // The range loses its end and gains a reference: unbalanced.
    const document = edited(
      previous,
      { type: "replace", at: 1, block: anchor("commentReference") },
      false,
    );
    const stories = changedStoriesOf(previous, document);
    const verdict = unsupportedSeedReasonWholeDocument(document, stories);
    expect(verdict).not.toBeNull();
    expect(unsupportedEditReason({ previous, document, stories })).toBe(
      verdict,
    );
  });
});
