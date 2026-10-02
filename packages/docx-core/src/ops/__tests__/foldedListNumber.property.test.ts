/**
 * The captures a reader makes of the `LISTNUM` fields a list marker draws are
 * paragraph content, so the operations that cut and join paragraphs own them
 * the way they own any other item: a split leaves each capture in exactly one
 * half, a join carries every capture of both paragraphs into the survivor, and
 * typed text moves none. After any sequence of the three, the story holds the
 * captures it started with, once each and in the same order.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document, Paragraph, ParagraphContent } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { paragraphLength } from "../offsets";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  SPLIT_HALVES,
} from "../types";

setDefaultTimeout(propertyTestTimeout(30_000));

type ItemKind = "text" | "field" | "fieldAndTab";

const ITEM_KINDS = ["text", "text", "field", "fieldAndTab"] as const satisfies readonly ItemKind[];

const paragraphKindsArbitrary = fc.array(
  fc.array(fc.constantFrom(...ITEM_KINDS), { minLength: 1, maxLength: 5 }),
  { minLength: 1, maxLength: 3 },
);

const capture = (kind: "field" | "tab", serial: number): ParagraphContent => ({
  type: "preservedInline",
  xml: `<w:r data-serial="${serial}"/>`,
  text: "",
  foldedListNumber:
    kind === "field"
      ? {
          kind,
          field: {
            type: "complexField",
            instruction: " LISTNUM ",
            fieldType: "LISTNUM",
            fieldCode: [],
            fieldResult: [{ type: "run", content: [{ type: "text", text: `(${serial})` }] }],
          },
        }
      : { kind, run: { type: "run", content: [{ type: "tab" }] } },
});

/** A story of paragraphs whose captures each carry markup of their own. */
const documentOf = (kinds: readonly (readonly ItemKind[])[]): Document => {
  let serial = 0;
  const content = kinds.map((items, index): Paragraph => {
    const inline: ParagraphContent[] = [];
    for (const item of items) {
      if (item === "text") {
        // One run per stretch of text: two runs alike side by side are one.
        if (inline.at(-1)?.type !== "run") {
          inline.push({ type: "run", content: [{ type: "text", text: "ab" }] });
        }
        continue;
      }
      inline.push(capture("field", serial++));
      if (item === "fieldAndTab") {
        inline.push(capture("tab", serial++));
      }
    }
    return { type: "paragraph", paraId: `0000000${index + 1}`, content: inline };
  });
  return { package: { document: { content } } };
};

const paragraphsOf = (document: Document): Paragraph[] =>
  document.package.document.content.flatMap((block) => (block.type === "paragraph" ? [block] : []));

const capturesOf = (document: Document): unknown[] =>
  paragraphsOf(document).flatMap((paragraph) =>
    paragraph.content.flatMap((item) =>
      item.type === "preservedInline" && item.foldedListNumber !== undefined
        ? [[item.foldedListNumber.kind, item.xml]]
        : [],
    ),
  );

const fractionArbitrary = fc.nat({ max: 1000 }).map((thousandths) => thousandths / 1000);

const STEP_KINDS = ["split", "join", "type"] as const;

const stepArbitrary = fc.record({
  kind: fc.constantFrom(...STEP_KINDS),
  paragraph: fractionArbitrary,
  offset: fractionArbitrary,
  zeroWidthBefore: fc.option(fc.nat({ max: 3 }), { nil: undefined }),
  half: fc.constantFrom(SPLIT_HALVES.FIRST, SPLIT_HALVES.SECOND),
});

type Step = {
  kind: (typeof STEP_KINDS)[number];
  paragraph: number;
  offset: number;
  zeroWidthBefore: number | undefined;
  half: (typeof SPLIT_HALVES)[keyof typeof SPLIT_HALVES];
};

const opFor = (document: Document, step: Step, serial: number): DocumentOp | undefined => {
  const paragraphs = paragraphsOf(document);
  const index = Math.min(paragraphs.length - 1, Math.floor(step.paragraph * paragraphs.length));
  const paragraph = paragraphs[index];
  if (!paragraph?.paraId) {
    return undefined;
  }
  const offset = Math.round(step.offset * paragraphLength(paragraph));
  const at =
    step.zeroWidthBefore === undefined
      ? { story: OP_STORIES.MAIN, blockId: paragraph.paraId, offset }
      : {
          story: OP_STORIES.MAIN,
          blockId: paragraph.paraId,
          offset,
          zeroWidthBefore: step.zeroWidthBefore,
        };
  switch (step.kind) {
    case "split":
      return {
        type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
        at,
        newBlockId: (0x10_00 + serial).toString(16).toUpperCase().padStart(8, "0"),
        newHalf: step.half,
      };
    case "join": {
      const next = paragraphs[index + 1];
      if (!next?.paraId) {
        return undefined;
      }
      return {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId: paragraph.paraId,
        nextBlockId: next.paraId,
        survivor: step.half,
      };
    }
    case "type":
      return { type: DOCUMENT_OP_TYPES.INSERT_TEXT, at, text: "x", runProps: INHERIT_RUN_PROPS };
    default: {
      const unhandled: never = step.kind;
      return unhandled;
    }
  }
};

describe("the captures of folded LISTNUM fields under split, join and typing", () => {
  test("every capture stays in the story once, in the order it started in", () => {
    assertProperty(
      fc.property(
        paragraphKindsArbitrary,
        fc.array(stepArbitrary, { minLength: 1, maxLength: 10 }),
        (kinds, steps) => {
          let document = documentOf(kinds);
          const original = capturesOf(document);

          for (const [serial, step] of steps.entries()) {
            const op = opFor(document, step, serial);
            if (op === undefined) {
              continue;
            }
            const result = applyDocumentOp(document, op);
            // A refused operation leaves the document as it was.
            if (result.isErr()) {
              continue;
            }
            const paragraphsBefore = paragraphsOf(document).length;
            document = result.value.document;
            expect(capturesOf(document)).toEqual(original);
            if (step.kind === "split") {
              expect(paragraphsOf(document)).toHaveLength(paragraphsBefore + 1);
            }
            if (step.kind === "join") {
              expect(paragraphsOf(document)).toHaveLength(paragraphsBefore - 1);
            }
          }
        },
      ),
      { numRuns: 300 },
    );
  });

  test("a split puts a capture in one half and a join brings it back", () => {
    const document = documentOf([["text", "fieldAndTab", "text"]]);
    const split = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 2 },
      newBlockId: "0000ABCD",
    });
    if (split.isErr()) {
      throw split.error;
    }

    const halves = paragraphsOf(split.value.document);
    expect(halves).toHaveLength(2);
    const held = halves.map(
      (paragraph) => capturesOf({ package: { document: { content: [paragraph] } } }).length,
    );
    expect(held.reduce((sum, count) => sum + count, 0)).toBe(2);
    expect(capturesOf(split.value.document)).toEqual(capturesOf(document));

    const [first, second] = halves;
    if (!first?.paraId || !second?.paraId) {
      throw new Error("Both halves carry an id");
    }
    const joined = applyDocumentOp(split.value.document, {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: first.paraId,
      nextBlockId: second.paraId,
    });
    if (joined.isErr()) {
      throw joined.error;
    }
    expect(paragraphsOf(joined.value.document)).toHaveLength(1);
    expect(capturesOf(joined.value.document)).toEqual(capturesOf(document));
  });
});
