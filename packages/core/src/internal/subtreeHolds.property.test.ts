/**
 * Walks that skip the subtrees holding none of the nodes they need, decided
 * once per node, answer what the walks over every node answer.
 *
 * Along a chain of edits that share every block they do not touch, the page
 * break run index gives every node the position the full walk gives it, and
 * numbered REF fields resolve to the same values.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { Fragment, type Node as PMNode } from "prosemirror-model";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import {
  resolveNumberedRefFields,
  resolveNumberedRefFieldsVisitingEveryNode,
} from "../prosemirror/numberedRefFields";
import { schema } from "../prosemirror/schema";
import type { Paragraph, ParagraphContent } from "../types/document";
import {
  buildPageBreakRunDescendantIndex,
  buildPageBreakRunDescendantIndexVisitingEveryNode,
} from "./pageBreakRunDescendantIndex";

// Synchronous generated edit chains use a 15-second base budget.
setDefaultTimeout(propertyTestTimeout(15_000));

const textRun = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const BOOKMARKS = ["clause-a", "clause-b"] as const;

/** A numbered paragraph, maybe holding a bookmark or a numbered REF field. */
const numberedParagraph = (
  level: number,
  bookmark: string | undefined,
  reference: { bookmark: string; displayText: string } | undefined,
): Paragraph => ({
  type: "paragraph",
  formatting: { numPr: { kind: "reference", numId: 5, ilvl: level } },
  content: [
    ...(bookmark === undefined
      ? []
      : [{ type: "bookmarkStart" as const, id: level + 1, name: bookmark }]),
    textRun(`item ${String(level)}`),
    ...(bookmark === undefined
      ? []
      : [{ type: "bookmarkEnd" as const, id: level + 1 }]),
    ...(reference === undefined
      ? []
      : [
          {
            type: "complexField" as const,
            instruction: `REF ${reference.bookmark} \\r \\h`,
            fieldType: "REF",
            fieldCode: [],
            fieldResult: [textRun(reference.displayText)],
          },
        ]),
  ],
  listRendering: {
    marker: `%${String(level + 1)}.`,
    level,
    numId: 5,
    abstractNumId: 7,
    isBullet: false,
    numFmt: "decimal",
    levelNumFmts: ["decimal", "decimal", "decimal"],
    levelStarts: [1, 1, 1],
  },
});

const paragraphOf = fc
  .record({
    level: fc.integer({ min: 0, max: 2 }),
    bookmark: fc.option(fc.constantFrom(...BOOKMARKS), { nil: undefined }),
    reference: fc.option(
      fc.record({
        bookmark: fc.constantFrom(...BOOKMARKS),
        displayText: fc.constantFrom("1", "2", "1.1", "(a)"),
      }),
      { nil: undefined },
    ),
    pageBreak: fc.boolean(),
  })
  .map(({ level, bookmark, reference, pageBreak }) => {
    const converted = toProseDoc({
      package: {
        document: { content: [numberedParagraph(level, bookmark, reference)] },
      },
    }).firstChild;
    if (converted === null) throw new TypeError("No paragraph converted.");
    return pageBreak
      ? converted.copy(
          converted.content.append(Fragment.from(schema.node("pageBreakRun"))),
        )
      : converted;
  });

const blockOf: fc.Arbitrary<PMNode> = fc.oneof(
  { weight: 3, arbitrary: paragraphOf },
  {
    weight: 1,
    arbitrary: fc
      .array(paragraphOf, { minLength: 1, maxLength: 2 })
      .map((cell) =>
        schema.node("table", null, [
          schema.node("tableRow", null, [schema.node("tableCell", null, cell)]),
        ]),
      ),
  },
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

/** Every node of the document, the root first, in document order. */
const nodesOf = (doc: PMNode): PMNode[] => {
  const nodes = [doc];
  doc.descendants((node) => {
    nodes.push(node);
    return true;
  });
  return nodes;
};

const sameAnswers = (doc: PMNode): void => {
  const nodes = nodesOf(doc);
  const skipping = buildPageBreakRunDescendantIndex(doc);
  const full = buildPageBreakRunDescendantIndexVisitingEveryNode(doc);
  expect(nodes.map((node) => skipping.firstPageBreakRunPosition(node))).toEqual(
    nodes.map((node) => full.firstPageBreakRunPosition(node)),
  );
  const resolved = resolveNumberedRefFields(doc);
  const reference = resolveNumberedRefFieldsVisitingEveryNode(doc);
  expect(nodes.map((node) => resolved.get(node))).toEqual(
    nodes.map((node) => reference.get(node)),
  );
  expect(resolved.size).toBe(reference.size);
};

describe("walks skipping subtrees that hold nothing they need", () => {
  test("answer what the walks over every node answer", () => {
    assertProperty(
      fc.property(
        fc.array(blockOf, { minLength: 1, maxLength: 6 }),
        fc.array(editOf, { minLength: 1, maxLength: 8 }),
        (blocks, edits) => {
          let doc = schema.node("doc", null, blocks);
          sameAnswers(doc);
          for (const edit of edits) {
            doc = edited(doc, edit);
            sameAnswers(doc);
          }
        },
      ),
      { id: "subtree skipping walks equal full walks", numRuns: 300 },
    );
  });
});
