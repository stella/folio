/**
 * Bookmark pairing skips the subtrees that hold no boundary, decided once per
 * node. Along a chain of edits that share every subtree they do not touch,
 * the issues it reports equal the pass that walks every node.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { Fragment, type Node as PMNode } from "prosemirror-model";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { schema } from "./schema";
import { bookmarkIssues, bookmarkIssuesWalkingEveryNode } from "./validation";

// Synchronous generated edit chains use a 15-second base budget.
setDefaultTimeout(propertyTestTimeout(15_000));

const ids = fc.integer({ min: 1, max: 3 });

const inlineOf: fc.Arbitrary<PMNode> = fc.oneof(
  {
    weight: 3,
    arbitrary: fc.constantFrom("a", "bc").map((text) => schema.text(text)),
  },
  {
    weight: 1,
    arbitrary: fc
      .record({ id: ids, start: fc.boolean() })
      .map(({ id, start }) =>
        schema.node(
          "bookmarkBoundary",
          start
            ? { type: "start", id, name: `b${String(id)}` }
            : { type: "end", id },
        ),
      ),
  },
);

const paragraphOf = fc
  .record({
    content: fc.array(inlineOf, { maxLength: 3 }),
    bookmarks: fc.array(ids, { maxLength: 1 }),
  })
  .map(({ content, bookmarks }) =>
    schema.node(
      "paragraph",
      bookmarks.length === 0
        ? null
        : {
            bookmarks: bookmarks.map((id) => ({ id, name: `p${String(id)}` })),
          },
      content,
    ),
  );

const markersOf = fc.array(
  fc.record({ id: ids, start: fc.boolean() }).map(({ id, start }) => ({
    index: 0,
    marker: start
      ? { type: "bookmarkStart" as const, id, name: `t${String(id)}` }
      : { type: "bookmarkEnd" as const, id },
  })),
  { maxLength: 1 },
);

const tableOf = fc
  .record({
    table: markersOf,
    row: markersOf,
    cell: fc.array(paragraphOf, { minLength: 1, maxLength: 2 }),
  })
  .map(({ table, row, cell }) =>
    schema.node("table", table.length === 0 ? null : { _bookmarks: table }, [
      schema.node("tableRow", row.length === 0 ? null : { _bookmarks: row }, [
        schema.node("tableCell", null, cell),
      ]),
    ]),
  );

const blockOf = fc.oneof(
  { weight: 3, arbitrary: paragraphOf },
  { weight: 1, arbitrary: tableOf },
);

type Edit =
  | { type: "replace"; at: number; block: PMNode }
  | { type: "insert"; at: number; block: PMNode }
  | { type: "delete"; at: number };

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
);

/** A new document sharing every block the edit does not touch. */
const edited = (doc: PMNode, edit: Edit): PMNode => {
  const blocks: PMNode[] = [];
  doc.forEach((block) => blocks.push(block));
  const at = edit.at % blocks.length;
  if (edit.type === "replace") blocks.splice(at, 1, edit.block);
  else if (edit.type === "insert") blocks.splice(at, 0, edit.block);
  else if (blocks.length > 1) blocks.splice(at, 1);
  return doc.copy(Fragment.from(blocks));
};

describe("bookmark validation skip", () => {
  test("reports what the pass over every node reports", () => {
    fc.assert(
      fc.property(
        fc.array(blockOf, { minLength: 1, maxLength: 5 }),
        fc.array(editOf, { minLength: 1, maxLength: 8 }),
        (blocks, edits) => {
          let doc = schema.node("doc", null, blocks);
          expect(bookmarkIssues(doc)).toEqual(
            bookmarkIssuesWalkingEveryNode(doc),
          );
          for (const edit of edits) {
            doc = edited(doc, edit);
            expect(bookmarkIssues(doc)).toEqual(
              bookmarkIssuesWalkingEveryNode(doc),
            );
          }
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });
});
