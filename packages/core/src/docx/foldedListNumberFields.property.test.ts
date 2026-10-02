/**
 * The two halves of the list-number fold, as laws.
 *
 * The reader's half changes what two kinds of item are and nothing about
 * where anything stands: each `LISTNUM` field that opens a paragraph, and the
 * tab after it, becomes a capture in the same place, and every other item is
 * the same object at the same index.
 *
 * The other half is the rule that says which captures may stay hidden once
 * the content has been edited. It never hides a capture behind something the
 * line shows, never shows in the marker a field it does not hide, and is done
 * after one pass.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type {
  ComplexField,
  Paragraph,
  ParagraphContent,
  PreservedInline,
  Run,
} from "../types/document";
import {
  foldedListNumberOf,
  foldListNumberFields,
  isFoldedListNumberCapture,
  isTabOnlyRun,
  type ListNumberFoldItem,
  normalizeFoldedListNumbers,
  planListNumberFold,
  unfoldedListNumberContent,
} from "./foldedListNumberFields";
import { createEmptyDocument } from "../utils/createDocument";
import { createDocx } from "./rezip";
import { serializeParagraph } from "./serializer/paragraphSerializer";

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
    { weight: 2, arbitrary: fc.constantFrom("a", "bc", "%1", "x y", "50%").map(textRun) },
    { weight: 3, arbitrary: fc.constant(null).map(tabRun) },
    {
      weight: 5,
      arbitrary: fc
        .tuple(
          fc.constantFrom(" LISTNUM ", "LISTNUM \\l 3", " listnum LegalDefault "),
          fc.constantFrom("(a)", "(ii)", "%", ""),
        )
        .map(([instruction, result]): ParagraphContent => field(instruction, result)),
    },
    { weight: 1, arbitrary: fc.constant(null).map((): ParagraphContent => field(" PAGE ", "3")) },
    {
      weight: 4,
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

const markupOf = (content: readonly ParagraphContent[]) => {
  const markup = new Map<ParagraphContent, string>();
  for (const [index, item] of content.entries()) {
    markup.set(item, `<w:r data-index="${index}"/>`);
  }
  return (item: ParagraphContent): string | undefined => markup.get(item);
};

describe("the reader's fold of the LISTNUM fields that open a paragraph", () => {
  test("every item keeps its place, and a capture stands for the item that was there", () => {
    assertProperty(
      fc.property(contentArbitrary, (content) => {
        const source = markupOf(content);

        const fold = foldListNumberFields(content, source);

        expect(fold.content).toHaveLength(content.length);
        for (const [index, item] of content.entries()) {
          const folded = fold.content[index];
          if (folded?.type === "preservedInline") {
            expect(unfoldedListNumberContent(folded)).toBe(item);
            expect(folded.xml).toBe(source(item) ?? "");
            expect(folded.text).toBe("");
          } else {
            expect(folded).toBe(item);
          }
        }
      }),
      { numRuns: 300 },
    );
  });

  test("the captures are the fields ahead of anything shown, and the tabs that follow them", () => {
    assertProperty(
      fc.property(contentArbitrary, (content) => {
        const fold = foldListNumberFields(content, markupOf(content));

        expect(fold.fieldCount).toBe(content.filter(isListNumberField).length);
        const cached: string[] = [];
        let opening = true;
        let afterField = false;
        for (const [index, item] of content.entries()) {
          const folded = fold.content[index];
          const kind = folded ? foldedListNumberOf(folded)?.kind : undefined;
          if (!opening) {
            expect(kind).toBeUndefined();
          } else if (isListNumberField(item)) {
            expect(kind).toBe("field");
            if (item.type === "complexField") {
              cached.push(
                ...item.fieldResult
                  .flatMap((run) => run.content)
                  .flatMap((piece) => (piece.type === "text" ? [piece.text] : [])),
              );
            }
            afterField = true;
          } else if (isMarker(item)) {
            // A marker neither is folded nor ends the wait for the tab.
            expect(kind).toBeUndefined();
          } else if (afterField && isTabOnlyRun(item)) {
            expect(kind).toBe("tab");
            afterField = false;
          } else {
            expect(kind).toBeUndefined();
            opening = false;
          }
        }
        expect(fold.cached).toEqual(cached);
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

const foldItemArbitrary: fc.Arbitrary<ListNumberFoldItem> = fc.oneof(
  {
    weight: 3,
    arbitrary: fc
      .constantFrom("(a)", "(b)", "")
      .map((cached): ListNumberFoldItem => ({ kind: "field", cached })),
  },
  { weight: 2, arbitrary: fc.constant<ListNumberFoldItem>({ kind: "tab" }) },
  { weight: 2, arbitrary: fc.constant<ListNumberFoldItem>({ kind: "hidden" }) },
  { weight: 2, arbitrary: fc.constant<ListNumberFoldItem>({ kind: "shown" }) },
);

const itemsArbitrary = fc.array(foldItemArbitrary, { maxLength: 10 });

const isCapture = (item: ListNumberFoldItem | undefined): boolean =>
  item?.kind === "field" || item?.kind === "tab";

/** The items as they stand once the plan is carried out: a capture not hidden is shown. */
const carriedOut = (
  items: readonly ListNumberFoldItem[],
  plan: ReturnType<typeof planListNumberFold>,
): ListNumberFoldItem[] =>
  plan.order.flatMap((index): ListNumberFoldItem[] => {
    const item = items[index];
    if (!item) {
      return [];
    }
    return isCapture(item) && !plan.hidden.has(index) ? [{ kind: "shown" }] : [item];
  });

describe("the rule for which captures stay hidden", () => {
  test("a marker that shows no fields hides none", () => {
    assertProperty(
      fc.property(itemsArbitrary, (items) => {
        const plan = planListNumberFold(items, false);

        expect(plan.order).toEqual(items.map((_, index) => index));
        expect(plan.hidden.size).toBe(0);
        expect(plan.suffix).toBeUndefined();
      }),
      { numRuns: 200 },
    );
  });

  test("it reorders nothing but one stretch of captures, and keeps every item once", () => {
    assertProperty(
      fc.property(itemsArbitrary, (items) => {
        const plan = planListNumberFold(items, true);

        expect(plan.order.toSorted((a, b) => a - b)).toEqual(items.map((_, index) => index));
        // What the line shows never changes order.
        const shown = plan.order.filter((index) => items[index]?.kind === "shown");
        expect(shown).toEqual(shown.toSorted((a, b) => a - b));
        const captures = plan.order.filter((index) => isCapture(items[index]));
        expect(captures).toEqual(captures.toSorted((a, b) => a - b));
      }),
      { numRuns: 300 },
    );
  });

  test("a hidden capture has nothing shown ahead of it, and the marker shows exactly the hidden fields", () => {
    assertProperty(
      fc.property(itemsArbitrary, (items) => {
        const plan = planListNumberFold(items, true);
        const after = carriedOut(items, plan);

        const lastHidden = after.findLastIndex((item) => isCapture(item));
        expect(after.slice(0, lastHidden + 1).some((item) => item.kind === "shown")).toBe(false);
        for (const index of plan.hidden) {
          expect(isCapture(items[index])).toBe(true);
        }
        const cached = after.flatMap((item) =>
          item.kind === "field" && item.cached !== "" ? [item.cached] : [],
        );
        expect(plan.suffix).toBe(cached.length === 0 ? undefined : cached.join(" "));
        // A marker with nothing to show hides nothing at all.
        if (plan.suffix === undefined) {
          expect(after.some((item) => isCapture(item))).toBe(false);
        }
        // A tab is hidden only right behind its field.
        for (const [index, item] of after.entries()) {
          if (item.kind === "tab") {
            const before = after.slice(0, index).findLast((earlier) => earlier.kind !== "hidden");
            expect(before?.kind).toBe("field");
          }
        }
      }),
      { numRuns: 300 },
    );
  });

  test("one pass is enough: the items it leaves are already in the form it allows", () => {
    assertProperty(
      fc.property(itemsArbitrary, (items) => {
        const first = planListNumberFold(items, true);
        const after = carriedOut(items, first);

        const second = planListNumberFold(after, first.suffix !== undefined);

        expect(second.order).toEqual(after.map((_, index) => index));
        expect(second.suffix).toBe(first.suffix);
        expect(carriedOut(after, second)).toEqual(after);
      }),
      { numRuns: 300 },
    );
  });

  test("text put ahead of the captures that open a paragraph ends up behind them", () => {
    const items: ListNumberFoldItem[] = [
      { kind: "shown" },
      { kind: "field", cached: "(a)" },
      { kind: "hidden" },
      { kind: "tab" },
      { kind: "hidden" },
      { kind: "shown" },
    ];

    const plan = planListNumberFold(items, true);

    expect(plan.order).toEqual([1, 2, 3, 0, 4, 5]);
    expect([...plan.hidden]).toEqual([1, 3]);
    expect(plan.suffix).toBe("(a)");
  });

  test("a capture behind the paragraph's own text is shown, and the one that opens it stays", () => {
    const items: ListNumberFoldItem[] = [
      { kind: "field", cached: "(a)" },
      { kind: "tab" },
      { kind: "shown" },
      { kind: "field", cached: "(b)" },
      { kind: "tab" },
    ];

    const plan = planListNumberFold(items, true);

    expect(plan.order).toEqual([0, 1, 2, 3, 4]);
    expect([...plan.hidden]).toEqual([0, 1]);
    expect(plan.suffix).toBe("(a)");
  });
});

describe("a paragraph of the model brought to the form the fold allows", () => {
  const listNumber = field(" LISTNUM ", "(a)");
  const tab = tabRun();
  const captures = (): ParagraphContent[] =>
    foldListNumberFields([listNumber, tab], (item) => `<w:r data-type="${item.type}"/>`).content;

  const paragraph = (content: ParagraphContent[], marker: string | undefined): Paragraph => ({
    type: "paragraph",
    content,
    ...(marker === undefined
      ? {}
      : {
          listRendering: {
            marker,
            markerTemplate: "%1.%2",
            level: 1,
            numId: 4,
            isBullet: false,
          },
        }),
  });

  test("captures its marker shows stay as they are", () => {
    const content = [...captures(), textRun("Body")];
    const numbered = paragraph(content, "7.1\t(a)");

    normalizeFoldedListNumbers(numbered);

    expect(numbered.content).toEqual(content);
    expect(numbered.listRendering?.marker).toBe("7.1\t(a)");
  });

  test("a paragraph with no numbering shows the field and the tab", () => {
    const plain = paragraph([...captures(), textRun("Body")], undefined);

    normalizeFoldedListNumbers(plain);

    expect(plain.content).toEqual([listNumber, tab, textRun("Body")]);
  });

  test("a marker that no longer shows the field puts it on the line", () => {
    const moved = paragraph([...captures(), textRun("Body")], "(a)");

    normalizeFoldedListNumbers(moved);

    expect(moved.content).toEqual([listNumber, tab, textRun("Body")]);
    expect(moved.listRendering?.marker).toBe("(a)");
  });

  test("a marker left showing a field the paragraph no longer holds stops showing it", () => {
    const emptied = paragraph([textRun("Body")], "7.1\t(a)");

    normalizeFoldedListNumbers(emptied);

    expect(emptied.listRendering?.marker).toBe("7.1");
  });

  test("text ahead of the captures goes behind them", () => {
    const typed = paragraph([textRun("X"), ...captures(), textRun("Body")], "7.1\t(a)");

    normalizeFoldedListNumbers(typed);

    expect(
      typed.content.map((item) => (isFoldedListNumberCapture(item) ? "capture" : item)),
    ).toEqual(["capture", "capture", textRun("X"), textRun("Body")]);
    expect(typed.listRendering?.marker).toBe("7.1\t(a)");
  });

  test("a capture inside a tracked deletion is written as a deleted field", () => {
    const deleted = paragraph(
      [
        {
          type: "deletion",
          info: { id: 1, author: "Reviewer" },
          content: captures().flatMap((item) => (item.type === "preservedInline" ? [item] : [])),
        },
        textRun("Body"),
      ],
      "7.1\t(a)",
    );

    normalizeFoldedListNumbers(deleted);

    const [wrapper] = deleted.content;
    expect(wrapper?.type === "deletion" ? wrapper.content : []).toEqual([listNumber, tab]);
    // Nothing is hidden any more, so the marker shows no field.
    expect(deleted.listRendering?.marker).toBe("7.1");
    const xml = serializeParagraph(deleted);
    expect(xml).toContain("<w:delInstrText");
    expect(xml).toMatch(/<w:delText[^>]*>\(a\)<\/w:delText>/u);
    expect(xml).not.toContain("<w:instrText");
    expect(xml).not.toContain("data-type");
  });

  const rawCapture = (xml: string): PreservedInline => ({
    type: "preservedInline",
    xml,
    text: "",
    foldedListNumber: { kind: "field", field: listNumber },
  });

  const FIELD_MARKUP =
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> LISTNUM </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>(a)</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

  const deleting = (capture: PreservedInline): Paragraph => ({
    type: "paragraph",
    content: [{ type: "deletion", info: { id: 1, author: "Reviewer" }, content: [capture] }],
  });

  test("the writer respells a capture it finds inside a removal, whoever left it there", () => {
    const xml = serializeParagraph(deleting(rawCapture(FIELD_MARKUP)));

    expect(xml).toContain("<w:delInstrText> LISTNUM </w:delInstrText>");
    expect(xml).toContain("<w:delText>(a)</w:delText>");
    expect(xml).not.toContain("<w:instrText");
    expect(xml).not.toMatch(/<w:t[\s>]/u);
  });

  test("the writer refuses a capture it cannot respell rather than write live text inside a removal", () => {
    const withDrawing = rawCapture(
      `<w:r><w:drawing/></w:r><w:r><w:instrText> LISTNUM </w:instrText></w:r>`,
    );

    expect(() => serializeParagraph(deleting(withDrawing))).toThrow(
      /cannot be written as deleted content/u,
    );
  });

  test("a save that meets such a capture is refused under the error's own tag", async () => {
    const document = createEmptyDocument();
    document.package.document.content = [
      deleting(
        rawCapture(`<w:r><w:drawing/></w:r><w:r><w:instrText> LISTNUM </w:instrText></w:r>`),
      ),
    ];

    // The same way a save reports other content a revision cannot hold.
    await expect(createDocx(document)).rejects.toMatchObject({
      _tag: "UnrepresentableTrackedCaptureError",
    });
  });

  test("a capture inside an insertion is written as it is", () => {
    const inserted: Paragraph = {
      type: "paragraph",
      content: [
        {
          type: "insertion",
          info: { id: 1, author: "Reviewer" },
          content: [rawCapture(FIELD_MARKUP)],
        },
      ],
    };

    expect(serializeParagraph(inserted)).toContain(FIELD_MARKUP);
  });

  test("a bullet hides nothing", () => {
    const bullet = paragraph([...captures(), textRun("Body")], "•\t(a)");
    if (bullet.listRendering) {
      bullet.listRendering.isBullet = true;
    }

    normalizeFoldedListNumbers(bullet);

    expect(bullet.content).toEqual([listNumber, tab, textRun("Body")]);
  });
});
