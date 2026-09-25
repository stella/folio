/**
 * Endnote areas paginate with the body (ECMA-376 §17.11).
 *
 * Measurement is a fixed-width provider so page breaks are predictable. Each
 * document is built in the model, serialized and parsed again, so the tests go
 * through the same package path an opened document does.
 */

import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import { buildDisplayList } from "../../display-list/build/buildDisplayList";
import type { DisplayHitRegion, DisplayPrimitive } from "../../display-list/types";
import { createDocx } from "../../docx/rezip";
import { layoutDocxHeadless, type HeadlessLayoutResult } from "../../headless-layout";
import { ptToPx } from "../../layout-engine/measure/measureHelpers";
import {
  getMeasureProvider,
  setMeasureProvider,
} from "../../layout-engine/measure/measureProvider";
import {
  NOTE_SEPARATOR_RULE_THICKNESS,
  NOTE_SEPARATOR_WIDTH_FRACTION,
  type Page,
  type ParagraphBlock,
} from "../../layout-engine/types";
import type { BlockContent, Document, Endnote, Paragraph } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";

let installedProvider = getMeasureProvider();

beforeEach(() => {
  installedProvider = getMeasureProvider();
  installFixedWidthProvider();
});

afterEach(() => {
  setMeasureProvider(installedProvider);
});

const FIXED_ADVANCE_RATIO = 0.5;

function installFixedWidthProvider(): void {
  const metricsOf = (fontSize: number | undefined) => {
    const px = ptToPx(fontSize ?? 11);
    return {
      fontSize: fontSize ?? 11,
      ascent: px * 0.8,
      descent: px * 0.2,
      fontBoxAscent: px * 0.9,
      fontBoxDescent: px * 0.25,
      lineHeight: px,
      fontFamily: "fixed",
      singleLineRatio: 1.15,
    };
  };
  const widthOf = (text: string, fontSize: number | undefined) =>
    [...text].length * ptToPx(fontSize ?? 11) * FIXED_ADVANCE_RATIO;
  setMeasureProvider({
    getFontMetrics: (style) => metricsOf(style.fontSize),
    measureTextWidth: (text, style) => widthOf(text, style.fontSize),
    measureText: (text, style) => {
      const metrics = metricsOf(style.fontSize);
      return {
        width: widthOf(text, style.fontSize),
        height: metrics.ascent + metrics.descent,
        ascent: metrics.ascent,
        descent: metrics.descent,
      };
    },
    measureRun: (text, style) => {
      const per = ptToPx(style.fontSize ?? 11) * FIXED_ADVANCE_RATIO;
      const charWidths = [...text].flatMap((char) => (char.length === 2 ? [per, 0] : [per]));
      return {
        width: charWidths.reduce((total, width) => total + width, 0),
        charWidths,
        metrics: metricsOf(style.fontSize),
      };
    },
  });
}

const textParagraph = (text: string): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const referencingParagraph = (text: string, endnoteId: number): Paragraph => ({
  type: "paragraph",
  content: [
    { type: "run", content: [{ type: "text", text }] },
    { type: "run", content: [{ type: "endnoteRef", id: endnoteId }] },
  ],
});

const endnote = (id: number, content: BlockContent[]): Endnote => ({
  type: "endnote",
  id,
  noteType: "normal",
  content,
});

const layOut = async (document: Document): Promise<HeadlessLayoutResult> => {
  const result = await layoutDocxHeadless(await createDocx(document), { pageGap: 0 });
  if (result.isErr()) {
    throw result.error;
  }
  return result.value;
};

const paragraphText = (block: ParagraphBlock): string =>
  block.runs.map((run) => (run.kind === "text" ? run.text : "")).join("");

type PlacedBlock = {
  page: Page;
  blockId: string;
  y: number;
  text: string;
  noteId: number | undefined;
  fromLine: number | undefined;
};

/** Every paragraph fragment in page order, with the text its block holds. */
const placedParagraphs = ({ layout, blockLookup }: HeadlessLayoutResult): PlacedBlock[] =>
  layout.pages.flatMap((page) =>
    page.fragments.flatMap((fragment) => {
      const entry = blockLookup.get(String(fragment.blockId));
      if (fragment.kind !== "paragraph" || entry?.block.kind !== "paragraph") {
        return [];
      }
      return [
        {
          page,
          blockId: String(fragment.blockId),
          y: fragment.y,
          text: paragraphText(entry.block),
          noteId: entry.noteStory?.noteId,
          fromLine: fragment.fromLine,
        },
      ];
    }),
  );

const LONG_TEXT = "Endnote words that wrap across several lines of the note area. ".repeat(4);

describe("endnote areas", () => {
  test("docEnd: the area follows the last body block, opened by the separator", async () => {
    const document = createEmptyDocument();
    document.package.document.content = [
      referencingParagraph("Body with a reference", 7),
      textParagraph("Last body paragraph"),
    ];
    document.package.endnotes = [endnote(7, [textParagraph("The endnote text.")])];

    const laidOut = await layOut(document);
    const placed = placedParagraphs(laidOut);

    expect(laidOut.unsupported).toEqual([]);
    expect(laidOut.layout.pages).toHaveLength(1);
    const lastBody = placed.findIndex(({ text }) => text === "Last body paragraph");
    const separatorIndex = placed.findIndex(({ blockId }) => blockId.includes("separator"));
    const noteIndex = placed.findIndex(({ noteId }) => noteId === 7);
    expect(lastBody).toBeGreaterThanOrEqual(0);
    expect(separatorIndex).toBe(lastBody + 1);
    expect(noteIndex).toBe(separatorIndex + 1);
    expect(placed[noteIndex]?.text).toContain("The endnote text.");

    // A `w:separator` rule spans part of the column, centred in its line.
    const page = laidOut.layout.pages[0];
    const separator = placed[separatorIndex];
    const rule = page?.noteSeparators?.at(0);
    expect(page?.noteSeparators).toHaveLength(1);
    expect(rule?.x).toBe(page?.margins.left);
    const columnWidth =
      (page?.size.w ?? 0) - (page?.margins.left ?? 0) - (page?.margins.right ?? 0);
    expect(rule?.width).toBeCloseTo(columnWidth * NOTE_SEPARATOR_WIDTH_FRACTION);
    expect(rule?.y).toBeGreaterThan(separator?.y ?? Number.POSITIVE_INFINITY);
    expect(rule?.y).toBeLessThan(placed[noteIndex]?.y ?? 0);
  });

  test("the note shows the same number as its body marker, in reference order", async () => {
    const document = createEmptyDocument();
    document.package.document.content = [
      referencingParagraph("First reference", 9),
      referencingParagraph("Second reference", 3),
    ];
    document.package.endnotes = [
      endnote(3, [textParagraph("Note three.")]),
      endnote(9, [textParagraph("Note nine.")]),
    ];

    const laidOut = await layOut(document);
    const blocks = [...laidOut.blockLookup.values()].map(({ block }) => block);
    const markerTexts = blocks.flatMap((block) =>
      block.kind === "paragraph"
        ? block.runs.flatMap((run) =>
            run.kind === "text" && run.endnoteRefId !== undefined ? [run.text] : [],
          )
        : [],
    );
    const notes = placedParagraphs(laidOut).filter(({ noteId }) => noteId !== undefined);

    // Endnotes number by first reference, in `w:numFmt` (lowerRoman by default).
    expect(markerTexts).toEqual(["i", "ii"]);
    expect(notes.map(({ noteId }) => noteId)).toEqual([9, 3]);
    expect(notes[0]?.text.startsWith("i")).toBe(true);
    expect(notes[0]?.text).toContain("Note nine.");
    expect(notes[1]?.text.startsWith("ii")).toBe(true);
    expect(notes[1]?.text).toContain("Note three.");
  });

  test("an area that overflows continues on a new page under the continuation separator", async () => {
    const document = createEmptyDocument();
    document.package.document.content = [referencingParagraph("Body", 1)];
    document.package.endnotes = [
      endnote(
        1,
        Array.from({ length: 30 }, (_, index) => textParagraph(`${String(index)} ${LONG_TEXT}`)),
      ),
    ];

    const laidOut = await layOut(document);
    const pages = laidOut.layout.pages;
    const placed = placedParagraphs(laidOut);

    expect(pages.length).toBeGreaterThan(1);
    for (const page of pages.slice(1)) {
      const first = placed.find((entry) => entry.page === page);
      expect(first?.blockId).toBe("endnote-continuation-separator");
      expect(first?.y).toBeCloseTo(page.margins.top);
      // A `w:continuationSeparator` rule spans the whole column.
      const rule = page.noteSeparators?.at(0);
      expect(page.noteSeparators).toHaveLength(1);
      expect(rule?.width).toBeCloseTo(page.size.w - page.margins.left - page.margins.right);
      // Note content resumes below the continuation separator.
      const resumed = placed.find((entry) => entry.page === page && entry.noteId === 1);
      expect(resumed?.y).toBeGreaterThan((rule?.y ?? 0) + NOTE_SEPARATOR_RULE_THICKNESS);
    }
    // The opening separator is on the first page only.
    expect(placed.filter(({ blockId }) => blockId === "endnote-area-0-separator")).toHaveLength(1);
    expect(placed.find(({ blockId }) => blockId === "endnote-area-0-separator")?.page).toBe(
      pages[0],
    );
  });

  test("sectEnd: each section's endnotes close that section", async () => {
    const document = createEmptyDocument();
    const finalSection = document.package.document.finalSectionProperties ?? {};
    document.package.document.finalSectionProperties = {
      ...finalSection,
      endnotePr: { position: "sectEnd" },
    };
    document.package.document.content = [
      referencingParagraph("Section one", 1),
      {
        ...textParagraph("End of section one"),
        sectionProperties: { ...finalSection, sectionStart: "continuous" },
      },
      referencingParagraph("Section two", 2),
    ];
    document.package.endnotes = [
      endnote(1, [textParagraph("Note for section one.")]),
      endnote(2, [textParagraph("Note for section two.")]),
    ];

    const placed = placedParagraphs(await layOut(document)).map(({ text, noteId }) =>
      noteId === undefined ? text : `note ${String(noteId)}`,
    );

    const order = [
      "End of section one",
      "note 1",
      placed.find((text) => text.startsWith("Section two")),
      "note 2",
    ].map((text) => placed.indexOf(text ?? ""));
    expect(order.every((index) => index >= 0)).toBe(true);
    expect(order).toEqual([...order].sort((left, right) => left - right));
  });

  test("sectEnd: a w:noEndnote section passes its endnotes to the next section", async () => {
    const document = createEmptyDocument();
    const finalSection = document.package.document.finalSectionProperties ?? {};
    document.package.document.finalSectionProperties = {
      ...finalSection,
      endnotePr: { position: "sectEnd" },
    };
    document.package.document.content = [
      referencingParagraph("Section one", 1),
      {
        ...textParagraph("End of section one"),
        sectionProperties: { ...finalSection, sectionStart: "continuous", noEndnote: true },
      },
      referencingParagraph("Section two", 2),
    ];
    document.package.endnotes = [
      endnote(1, [textParagraph("Note for section one.")]),
      endnote(2, [textParagraph("Note for section two.")]),
    ];

    const placed = placedParagraphs(await layOut(document)).map(({ text, noteId }) =>
      noteId === undefined ? text : `note ${String(noteId)}`,
    );

    const sectionTwo = placed.findIndex((text) => text.startsWith("Section two"));
    expect(sectionTwo).toBeGreaterThanOrEqual(0);
    expect(placed.indexOf("note 1")).toBeGreaterThan(sectionTwo);
    expect(placed.indexOf("note 2")).toBeGreaterThan(placed.indexOf("note 1"));
  });

  test("keeps the final paragraph after a table when the endnote area follows it", async () => {
    const buildDocument = (withEndnote: boolean): Document => {
      const document = createEmptyDocument();
      document.package.document.content = [
        withEndnote ? referencingParagraph("Body", 1) : textParagraph("Body"),
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [{ type: "tableCell", content: [textParagraph("Cell")] }],
            },
          ],
        },
        { type: "paragraph", content: [] },
      ];
      if (withEndnote) {
        document.package.endnotes = [endnote(1, [textParagraph("Note.")])];
      }
      return document;
    };

    const finalParagraphHeight = async (withEndnote: boolean): Promise<number | undefined> => {
      const { layout, blockLookup } = await layOut(buildDocument(withEndnote));
      const fragments = layout.pages.flatMap((page) => page.fragments);
      const tableIndex = fragments.findIndex(({ kind }) => kind === "table");
      const final = fragments[tableIndex + 1];
      const entry = final === undefined ? undefined : blockLookup.get(String(final.blockId));
      return entry?.measure.kind === "paragraph" ? entry.measure.totalHeight : undefined;
    };

    // The terminal anchor collapses only when it is the document's last block.
    expect(await finalParagraphHeight(false)).toBe(0);
    expect(await finalParagraphHeight(true)).toBeGreaterThan(0);
  });

  test("a document whose endnotes part has no reference lays out as before", async () => {
    const buildDocument = (withEndnotes: boolean): Document => {
      const document = createEmptyDocument();
      document.package.document.content = [
        textParagraph("Only body text"),
        textParagraph(LONG_TEXT),
      ];
      if (withEndnotes) {
        document.package.endnotes = [endnote(1, [textParagraph("Unreferenced.")])];
      }
      return document;
    };
    const geometry = ({ layout }: HeadlessLayoutResult) =>
      layout.pages.map((page) => ({
        separators: page.noteSeparators,
        fragments: page.fragments.map(({ kind, x, y, width }) => ({ kind, x, y, width })),
      }));

    const without = await layOut(buildDocument(false));
    const withPart = await layOut(buildDocument(true));

    expect(geometry(withPart)).toEqual(geometry(without));
    expect(withPart.layout.pages.every((page) => page.noteSeparators === undefined)).toBe(true);
  });

  test("painted endnote text is addressed by its note story, not by body positions", async () => {
    const document = createEmptyDocument();
    document.package.document.content = [referencingParagraph("Body", 4)];
    document.package.endnotes = [endnote(4, [textParagraph("Painted endnote.")])];

    const laidOut = await layOut(document);
    const noteFragments = laidOut.layout.pages
      .flatMap((page) => page.fragments)
      .filter(
        (fragment) => laidOut.blockLookup.get(String(fragment.blockId))?.noteStory !== undefined,
      );
    expect(noteFragments.length).toBeGreaterThan(0);
    for (const fragment of noteFragments) {
      expect(fragment.pmStart).toBeUndefined();
      expect(fragment.pmEnd).toBeUndefined();
    }

    const list = buildDisplayList({
      layout: laidOut.layout,
      blockLookup: laidOut.blockLookup,
      documentFeatures: laidOut.documentFeatures,
      ...laidOut.furniture,
    });
    const page = list.pages.at(0);
    const noteRegions: DisplayHitRegion[] = [];
    const walk = (regions: readonly DisplayHitRegion[]): void => {
      for (const region of regions) {
        if (region.kind === "note" && region.model?.story?.kind === "endnote") {
          noteRegions.push(region);
        }
        walk(region.children);
      }
    };
    walk(page?.regions ?? []);
    expect(noteRegions.map((region) => region.model?.story)).toContainEqual({
      kind: "endnote",
      id: 4,
    });

    // The separator rule is painted from the layout, as the column rules are.
    const rule = laidOut.layout.pages[0]?.noteSeparators?.at(0);
    const rects = (page?.primitives ?? []).filter(
      (primitive): primitive is Extract<DisplayPrimitive, { kind: "rect" }> =>
        primitive.kind === "rect",
    );
    expect(
      rects.some(
        ({ rect }) =>
          rect.xPx === rule?.x &&
          rect.yPx === rule?.y &&
          rect.widthPx === rule?.width &&
          rect.heightPx === NOTE_SEPARATOR_RULE_THICKNESS,
      ),
    ).toBe(true);
  });
});
