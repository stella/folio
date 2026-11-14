/**
 * The lock guard reads each block's controls once and reuses them while the
 * block stays the same object, as every block an operation does not name
 * does. Along a chain of edits that share blocks, the controls it reads equal
 * one walk of the whole document, and its verdict on each edit equals the
 * verdict on fresh copies, which share no object with anything read before.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import type { BlockContent, Paragraph } from "@stll/docx-core/model";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  controlsOf,
  controlsOfWholeDocument,
  lockViolationBetween,
} from "../canonical/canonicalLockGuard";
import type { Document } from "../types/document";

// Synchronous generated edit chains use a 15-second base budget.
setDefaultTimeout(propertyTestTimeout(15_000));

type Lock = "sdtLocked" | "contentLocked" | "sdtContentLocked" | "unlocked";

const locks = fc.constantFrom<Lock | undefined>(
  undefined,
  "sdtLocked",
  "contentLocked",
  "sdtContentLocked",
  "unlocked",
);
// Few ids, so controls share them and some have none.
const ids = fc.option(fc.integer({ min: 1, max: 4 }), { nil: undefined });
const words = fc.constantFrom("a", "b", "c");

const run = (text: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text }],
});

const paragraphOf = fc
  .array(
    fc.oneof(
      words.map(run),
      fc
        .record({ id: ids, lock: locks, text: words })
        .map(({ id, lock, text }) => ({
          type: "inlineSdt" as const,
          properties: {
            sdtType: "richText" as const,
            ...(id === undefined ? {} : { id }),
            ...(lock === undefined ? {} : { lock }),
          },
          content: [run(text)],
        })),
    ),
    { maxLength: 3 },
  )
  .map((content): Paragraph => ({ type: "paragraph", content }));

const blockOf: fc.Arbitrary<BlockContent> = fc.oneof(
  paragraphOf,
  fc
    .record({
      id: ids,
      lock: locks,
      inner: fc.array(paragraphOf, { minLength: 1, maxLength: 2 }),
    })
    .map(({ id, lock, inner }) => ({
      type: "blockSdt" as const,
      properties: {
        sdtType: "richText" as const,
        ...(id === undefined ? {} : { id }),
        ...(lock === undefined ? {} : { lock }),
      },
      content: inner,
    })),
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

describe("canonical lock guard", () => {
  test("reads the controls of shared blocks as the whole-document walk does, and judges edits as on fresh copies", () => {
    fc.assert(
      fc.property(
        fc.array(blockOf, { maxLength: 5 }),
        fc.array(blockOf, { maxLength: 2 }),
        fc.array(fc.tuple(editOf, fc.boolean()), {
          minLength: 1,
          maxLength: 8,
        }),
        (body, note, edits) => {
          let before = documentOf(body, note);
          expect(controlsOf(before)).toEqual(controlsOfWholeDocument(before));
          for (const [edit, inNote] of edits) {
            const after = edited(before, edit, inNote);
            expect(controlsOf(after)).toEqual(controlsOfWholeDocument(after));
            expect(lockViolationBetween(before, after)).toEqual(
              lockViolationBetween(
                structuredClone(before),
                structuredClone(after),
              ),
            );
            before = after;
          }
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });
});
