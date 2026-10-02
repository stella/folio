/**
 * Folding a paragraph's `LISTNUM` fields into its marker and writing them
 * back are inverses: whatever stood in the content, the fields and their tabs
 * return to where they were, and nothing else moves.
 *
 * The second property regroups the kept text into runs of its own before the
 * fields go back, which is what a pass through the editor does to it.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { ComplexField, Paragraph, ParagraphContent, Run } from "../types/document";
import {
  foldedListNumberFieldsOf,
  foldListNumberFields,
  withFoldedListNumberFields,
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

/** Ways a paragraph stops owning the fields its marker folded. */
const NUMBERING_CHANGES = ["level", "numId", "unnumbered", "unrecorded"] as const;

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

const NUM_ID = 4;
const LEVEL = 1;

const foldedParagraph = (content: readonly ParagraphContent[]): Paragraph => {
  const fold = foldListNumberFields(content);
  return {
    type: "paragraph",
    content: fold.content,
    listRendering: { marker: "%1.%2", level: LEVEL, numId: NUM_ID, isBullet: false },
    foldedListNumberFields: { numId: NUM_ID, level: LEVEL, fields: fold.fields },
  };
};

/** Adjacent text runs joined into one, as the editor regroups them. */
const regrouped = (content: readonly ParagraphContent[]): ParagraphContent[] => {
  const joined: ParagraphContent[] = [];
  for (const item of content) {
    const previous = joined.at(-1);
    const text = item.type === "run" ? item.content.at(0) : undefined;
    const previousText = previous?.type === "run" ? previous.content.at(0) : undefined;
    if (text?.type === "text" && previousText?.type === "text") {
      joined[joined.length - 1] = textRun(previousText.text + text.text);
    } else {
      joined.push(item);
    }
  }
  return joined;
};

/** The content one unit at a time, so where a run boundary falls does not show. */
const units = (content: readonly ParagraphContent[]): string[] =>
  content.flatMap((item) => {
    if (item.type === "run") {
      return item.content.flatMap((piece) =>
        piece.type === "text"
          ? [...piece.text].map((character) => `text:${character}`)
          : [piece.type],
      );
    }
    if (item.type === "complexField") {
      return [`field:${item.instruction}:${JSON.stringify(item.fieldResult)}`];
    }
    return ["id" in item ? `${item.type}:${String(item.id)}` : item.type];
  });

describe("folding LISTNUM fields out of a paragraph and writing them back", () => {
  test("takes out every LISTNUM field, the tab after it, and nothing else", () => {
    assertProperty(
      fc.property(contentArbitrary, (content) => {
        const fold = foldListNumberFields(content);
        const listNumberFields = content.filter(
          (item) => item.type === "complexField" && item.fieldType === "LISTNUM",
        );

        expect(fold.fields.map((folded) => folded.field)).toEqual(listNumberFields);
        expect(fold.content.some((item) => listNumberFields.includes(item))).toBe(false);
        expect(fold.content.length + fold.fields.length).toBe(
          content.length - fold.fields.filter(({ tab }) => tab !== undefined).length,
        );
        expect(fold.cached).toEqual(
          fold.fields.flatMap(({ field: folded }) =>
            folded.fieldResult.flatMap((run) =>
              run.content.flatMap((piece) => (piece.type === "text" ? [piece.text] : [])),
            ),
          ),
        );
      }),
      { numRuns: 300 },
    );
  });

  test("writes each field and tab back where it stood", () => {
    assertProperty(
      fc.property(contentArbitrary, (content) => {
        expect(withFoldedListNumberFields(foldedParagraph(content))).toEqual(content);
      }),
      { numRuns: 300 },
    );
  });

  test("writes each field back between the same characters once the text is regrouped", () => {
    assertProperty(
      fc.property(contentArbitrary, (content) => {
        const paragraph = foldedParagraph(content);
        paragraph.content = regrouped(paragraph.content);

        expect(units(withFoldedListNumberFields(paragraph))).toEqual(units(content));
      }),
      { numRuns: 300 },
    );
  });

  test("writes the fields back only under the numbering level that folded them", () => {
    assertProperty(
      fc.property(contentArbitrary, fc.constantFrom(...NUMBERING_CHANGES), (content, change) => {
        const paragraph = foldedParagraph(content);
        const rendering = paragraph.listRendering;
        if (!rendering) {
          throw new Error("The folded paragraph is numbered");
        }
        switch (change) {
          case "level":
            paragraph.listRendering = { ...rendering, level: LEVEL + 1 };
            break;
          case "numId":
            paragraph.listRendering = { ...rendering, numId: NUM_ID + 1 };
            break;
          case "unnumbered":
            delete paragraph.listRendering;
            break;
          case "unrecorded":
            delete paragraph.foldedListNumberFields;
            break;
          default: {
            const unhandled: never = change;
            throw new Error(`Unhandled change ${String(unhandled)}`);
          }
        }

        expect(foldedListNumberFieldsOf(paragraph)).toBeUndefined();
        expect(withFoldedListNumberFields(paragraph)).toBe(paragraph.content);
      }),
      { numRuns: 100 },
    );
  });
});

describe("writing folded LISTNUM fields back into edited content", () => {
  const listNumber = field(" LISTNUM ", "(a)");

  const paragraphWith = (
    content: ParagraphContent[],
    position: { offset: number; markersBefore: number; markersBeforeTab?: number },
  ): Paragraph => ({
    type: "paragraph",
    content,
    listRendering: { marker: "%1.%2\t(a)", level: LEVEL, numId: NUM_ID, isBullet: false },
    foldedListNumberFields: {
      numId: NUM_ID,
      level: LEVEL,
      fields: [{ field: listNumber, tab: tabRun(), ...position }],
    },
  });

  test("a field recorded past the end of the content goes at the end", () => {
    const paragraph = paragraphWith([textRun("ab")], { offset: 9, markersBefore: 0 });

    expect(withFoldedListNumberFields(paragraph)).toEqual([textRun("ab"), listNumber, tabRun()]);
  });

  test("a field whose markers are gone goes ahead of the text at its offset", () => {
    const paragraph = paragraphWith([textRun("ab"), textRun("cd")], {
      offset: 2,
      markersBefore: 2,
      markersBeforeTab: 1,
    });

    expect(withFoldedListNumberFields(paragraph)).toEqual([
      textRun("ab"),
      listNumber,
      tabRun(),
      textRun("cd"),
    ]);
  });

  test("a field recorded inside one run cuts the run and keeps its formatting on both halves", () => {
    const bold: Run = {
      type: "run",
      formatting: { bold: true },
      content: [{ type: "text", text: "abcd" }],
    };
    const paragraph = paragraphWith([bold], { offset: 1, markersBefore: 0 });

    expect(withFoldedListNumberFields(paragraph)).toEqual([
      { ...bold, content: [{ type: "text", text: "a" }] },
      listNumber,
      tabRun(),
      { ...bold, content: [{ type: "text", text: "bcd" }] },
    ]);
  });

  test("the paragraph's own content is left as it was", () => {
    const content = [textRun("abcd")];
    const paragraph = paragraphWith(content, { offset: 2, markersBefore: 0 });

    withFoldedListNumberFields(paragraph);

    expect(paragraph.content).toBe(content);
    expect(content).toEqual([textRun("abcd")]);
  });
});
