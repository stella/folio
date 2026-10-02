/**
 * The fold of a paragraph's `LISTNUM` fields changes what two kinds of item
 * are and nothing about where anything stands: every field, and the tab after
 * it, becomes the markup it was read from, in the same place, and every other
 * item is the same object at the same index.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { ComplexField, ParagraphContent, Run } from "../types/document";
import {
  foldListNumberFields,
  isFoldedListNumberCapture,
  isTabOnlyRun,
} from "./foldedListNumberFields";

setDefaultTimeout(propertyTestTimeout(30_000));

const textRun = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });

const tabRun = (): Run => ({ type: "run", content: [{ type: "tab" }] });

const field = (instruction: string, result: string): ComplexField => ({
  type: "complexField",
  instruction,
  fieldType: instruction.trim().toUpperCase().startsWith("LISTNUM") ? "LISTNUM" : "PAGE",
  fieldCode: [],
  fieldResult: result === "" ? [] : [textRun(result)],
});

const MARKER_TYPES = [
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
] as const;

const contentArbitrary: fc.Arbitrary<ParagraphContent[]> = fc.array(
  fc.oneof(
    { weight: 4, arbitrary: fc.constantFrom("a", "bc", "%1", "x y", "50%").map(textRun) },
    { weight: 3, arbitrary: fc.constant(null).map(tabRun) },
    {
      weight: 4,
      arbitrary: fc
        .tuple(
          fc.constantFrom(" LISTNUM ", "LISTNUM \\l 3", " listnum LegalDefault "),
          fc.constantFrom("(a)", "(ii)", "%", ""),
        )
        .map(([instruction, result]): ParagraphContent => field(instruction, result)),
    },
    { weight: 1, arbitrary: fc.constant(null).map((): ParagraphContent => field(" PAGE ", "3")) },
    {
      weight: 3,
      arbitrary: fc
        .tuple(fc.constantFrom(...MARKER_TYPES), fc.integer({ min: 1, max: 9 }))
        .map(([type, id]): ParagraphContent => {
          switch (type) {
            case "bookmarkStart":
              return { type, id, name: `mark${id}` };
            case "bookmarkEnd":
            case "commentRangeStart":
            case "commentRangeEnd":
              return { type, id };
            default: {
              const unhandled: never = type;
              return unhandled;
            }
          }
        }),
    },
  ),
  { maxLength: 12 },
);

const isListNumberField = (item: ParagraphContent): boolean =>
  item.type === "complexField" && item.fieldType === "LISTNUM";

const isMarker = (item: ParagraphContent): boolean =>
  MARKER_TYPES.some((type) => type === item.type);

/** Markup that names the item it stands for, so a capture can be traced back. */
const sourcesOf = (content: readonly ParagraphContent[]) => {
  const markup = new Map<ParagraphContent, string>();
  const items = new Map<string, ParagraphContent>();
  for (const [index, item] of content.entries()) {
    markup.set(item, `<w:r data-index="${index}"/>`);
    items.set(`<w:r data-index="${index}"/>`, item);
  }
  return {
    markupOf: (item: ParagraphContent): string | undefined => markup.get(item),
    itemOf: (xml: string): ParagraphContent | undefined => items.get(xml),
  };
};

describe("folding LISTNUM fields into captures of their own markup", () => {
  test("every item keeps its place, and a capture stands for the item that was there", () => {
    assertProperty(
      fc.property(contentArbitrary, (content) => {
        const sources = sourcesOf(content);

        const fold = foldListNumberFields(content, sources.markupOf);

        expect(fold.content).toHaveLength(content.length);
        const restored = fold.content.map((item) =>
          isFoldedListNumberCapture(item) ? sources.itemOf(item.xml) : item,
        );
        for (const [index, item] of content.entries()) {
          expect(restored[index]).toBe(item);
        }
      }),
      { numRuns: 300 },
    );
  });

  test("the captures are exactly the fields and the tabs that follow them", () => {
    assertProperty(
      fc.property(contentArbitrary, (content) => {
        const sources = sourcesOf(content);

        const fold = foldListNumberFields(content, sources.markupOf);

        expect(fold.fieldCount).toBe(content.filter(isListNumberField).length);
        let afterField = false;
        for (const [index, item] of content.entries()) {
          const folded = fold.content[index];
          if (!folded) {
            throw new Error(`The fold dropped the item at ${index}`);
          }
          const kind = isFoldedListNumberCapture(folded) ? folded.foldedListNumber : undefined;
          if (isListNumberField(item)) {
            expect(kind).toBe("field");
            expect(folded).toMatchObject({ text: "" });
            afterField = true;
          } else if (isMarker(item)) {
            // A marker neither is folded nor ends the wait for the tab.
            expect(kind).toBeUndefined();
          } else {
            expect(kind).toBe(afterField && isTabOnlyRun(item) ? "tab" : undefined);
            afterField = false;
          }
        }
      }),
      { numRuns: 300 },
    );
  });

  test("the cached display is each folded field's result text, in order", () => {
    assertProperty(
      fc.property(contentArbitrary, (content) => {
        const fold = foldListNumberFields(content, sourcesOf(content).markupOf);

        expect(fold.cached).toEqual(
          content.flatMap((item) =>
            item.type === "complexField" && item.fieldType === "LISTNUM"
              ? item.fieldResult.flatMap((run) =>
                  run.content.flatMap((piece) => (piece.type === "text" ? [piece.text] : [])),
                )
              : [],
          ),
        );
      }),
      { numRuns: 300 },
    );
  });

  test("an item with no markup to stand for it is left as it is", () => {
    assertProperty(
      fc.property(contentArbitrary, (content) => {
        const fold = foldListNumberFields(content, () => undefined);

        expect(fold.content).toHaveLength(content.length);
        for (const [index, item] of content.entries()) {
          expect(fold.content[index]).toBe(item);
        }
        expect(fold.cached).toEqual([]);
        expect(fold.fieldCount).toBe(content.filter(isListNumberField).length);
      }),
      { numRuns: 100 },
    );
  });
});
