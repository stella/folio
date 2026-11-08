/**
 * The seed check judges each block once and reuses the judgment while the
 * block stays the same object, as every block an operation does not name
 * does. Along a chain of edits that share blocks, its verdict equals one walk
 * of every story.
 */
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import type { BlockContent, Paragraph } from "@stll/docx-core/model";

import { propertyConfig } from "../../../../test/property-testing";
import { table } from "../docx/server/build";
import type { Document } from "../types/document";
import {
  unsupportedSeedReason,
  unsupportedSeedReasonWholeDocument,
} from "./canonicalSeedContent";

// A tab in a text node is one of the inline contents the session refuses.
const texts = fc.constantFrom("a", "b", "c\td");

const paragraphOf = fc
  .array(texts, { maxLength: 2 })
  .map((words): Paragraph => ({
    type: "paragraph",
    content: words.map((text) => ({
      type: "run",
      content: [{ type: "text", text }],
    })),
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

const documentOf = (body: BlockContent[], note: BlockContent[]): Document => ({
  package: {
    document: { content: body },
    footnotes: [{ type: "footnote", id: 1, content: note }],
  },
});

/** A new document sharing every block the edit does not touch. */
const edited = (document: Document, edit: Edit, inNote: boolean): Document => {
  const note = document.package.footnotes?.at(0)?.content ?? [];
  const blocks = [...(inNote ? note : document.package.document.content)];
  const at = blocks.length === 0 ? 0 : edit.at % blocks.length;
  if (edit.type === "replace" && blocks.length > 0)
    blocks.splice(at, 1, edit.block);
  else if (edit.type === "insert") blocks.splice(at, 0, edit.block);
  else if (edit.type === "delete") blocks.splice(at, 1);
  return inNote
    ? documentOf(document.package.document.content, blocks)
    : documentOf(blocks, note);
};

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
});
