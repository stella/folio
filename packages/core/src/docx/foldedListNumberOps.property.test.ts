/**
 * Document operations cut and join paragraphs without knowing the fold: a
 * capture is content to them, and goes where content goes. What a save owes
 * the file afterwards is decided by the fold's rule, applied to whatever the
 * operations left. So after any sequence of splits, joins and typed text, a
 * paragraph as the save writes it never hides a field anywhere but at its
 * start behind a marker that shows it, and every field is still written once.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document, Paragraph, ParagraphContent } from "../types/document";
import {
  fieldResultsInFile,
  fieldResultsShown,
  foldFaults,
  inlineTokens,
} from "./__tests__/listNumberFieldFixture";
import { normalizeFoldedListNumbers } from "./foldedListNumberFields";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  paragraphLength,
  SPLIT_HALVES,
} from "@stll/docx-core/ops";

setDefaultTimeout(propertyTestTimeout(30_000));

const resultOf = (serial: number): string => `(${serial})`;

/** The capture of a field the reader folded, with markup that reads as that field. */
const fieldCapture = (serial: number): ParagraphContent => ({
  type: "preservedInline",
  xml:
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> LISTNUM </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${resultOf(serial)}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
  text: "",
  foldedListNumber: {
    kind: "field",
    field: {
      type: "complexField",
      instruction: " LISTNUM ",
      fieldType: "LISTNUM",
      fieldCode: [],
      fieldResult: [{ type: "run", content: [{ type: "text", text: resultOf(serial) }] }],
    },
  },
});

const tabCapture = (): ParagraphContent => ({
  type: "preservedInline",
  xml: "<w:r><w:tab/></w:r>",
  text: "",
  foldedListNumber: { kind: "tab", run: { type: "run", content: [{ type: "tab" }] } },
});

type ParagraphShape = { fields: boolean[]; text: boolean };

const shapesArbitrary: fc.Arbitrary<ParagraphShape[]> = fc.array(
  fc.record({
    // One entry per field that opens the paragraph: whether a tab follows it.
    fields: fc.array(fc.boolean(), { maxLength: 2 }),
    text: fc.boolean(),
  }),
  { minLength: 1, maxLength: 3 },
);

/** Paragraphs as the reader leaves them: captures first, shown by the marker, then text. */
const documentOf = (
  shapes: readonly ParagraphShape[],
): { document: Document; results: string[] } => {
  let serial = 0;
  const results: string[] = [];
  const content = shapes.map(({ fields, text }, index): Paragraph => {
    const inline: ParagraphContent[] = [];
    const cached: string[] = [];
    for (const tab of fields) {
      serial += 1;
      cached.push(resultOf(serial));
      inline.push(fieldCapture(serial));
      if (tab) {
        inline.push(tabCapture());
      }
    }
    if (text) {
      inline.push({ type: "run", content: [{ type: "text", text: "abcd" }] });
    }
    results.push(...cached);
    const base = `${index + 1}.`;
    return {
      type: "paragraph",
      paraId: `0000000${index + 1}`,
      content: inline,
      listRendering: {
        marker: cached.length === 0 ? base : `${base}\t${cached.join(" ")}`,
        markerTemplate: "%1.",
        level: 0,
        numId: 1,
        isBullet: false,
      },
    };
  });
  return { document: { package: { document: { content } } }, results };
};

const paragraphsOf = (document: Document): Paragraph[] =>
  document.package.document.content.flatMap((block) => (block.type === "paragraph" ? [block] : []));

const fractionArbitrary = fc.nat({ max: 1000 }).map((thousandths) => thousandths / 1000);

const STEP_KINDS = ["split", "join", "type"] as const;

const stepArbitrary = fc.record({
  kind: fc.constantFrom(...STEP_KINDS),
  paragraph: fractionArbitrary,
  offset: fractionArbitrary,
  zeroWidthBefore: fc.option(fc.nat({ max: 4 }), { nil: undefined }),
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

/** The paragraphs as a save writes them: the fold's rule applied to a copy of each. */
const asSaved = (document: Document): Paragraph[] =>
  paragraphsOf(document).map((paragraph) => {
    const copy = structuredClone(paragraph);
    normalizeFoldedListNumbers(copy);
    return copy;
  });

const expectSavedForm = (document: Document, results: readonly string[]): void => {
  const written: string[] = [];
  for (const paragraph of asSaved(document)) {
    // Nothing hidden behind text, and the marker shows exactly what is hidden.
    expect(foldFaults(paragraph)).toEqual([]);
    const tokens = inlineTokens(serializeParagraph(paragraph));
    expect(fieldResultsShown(paragraph)).toBe(fieldResultsInFile(tokens));
    written.push(...tokens.filter((token) => /^text:\(\d+\)$/u.test(token)));
  }
  // Every field written once: none lost to a cut, none doubled by a join.
  expect(written.toSorted()).toEqual(results.map((result) => `text:${result}`).toSorted());
};

describe("the fold's rule after document operations", () => {
  test("no split, join or typed text leaves a field hidden where its marker does not show it", () => {
    assertProperty(
      fc.property(
        shapesArbitrary,
        fc.array(stepArbitrary, { minLength: 1, maxLength: 10 }),
        (shapes, steps) => {
          const built = documentOf(shapes);
          let document = built.document;
          expectSavedForm(document, built.results);

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
            document = result.value.document;
            expectSavedForm(document, built.results);
          }
        },
      ),
      { numRuns: 300 },
    );
  });

  test("a split behind the text moves the captures' paragraph on, and the save still hides them only at its start", () => {
    const { document, results } = documentOf([{ fields: [true], text: true }]);

    const split = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 2 },
      newBlockId: "0000ABCD",
    });
    if (split.isErr()) {
      throw split.error;
    }

    expect(paragraphsOf(split.value.document)).toHaveLength(2);
    expectSavedForm(split.value.document, results);
  });
});
