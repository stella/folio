/**
 * A page is not one document.
 *
 * Its body, each header and footer part, and each note are separate stories
 * with separate position spaces, and position 12 is a different character in
 * each of them. A run therefore carries the story its range belongs to, and an
 * editing surface routes a click into that story before it uses the number.
 *
 * Without the story there were only two options, and both lose: hand out the
 * number anyway and a click in a header edits the body, or drop it and a click
 * in a header edits nothing.
 */

import { afterEach, describe, expect, test } from "bun:test";

import { buildDisplayList } from "./buildDisplayList";
import type { DisplayGlyphRun, DisplayList, DisplayPrimitive } from "../types";
import { layoutDocxHeadless } from "../../headless-layout";
import { ptToPx } from "../../layout-engine/measure/measureHelpers";
import {
  getMeasureProvider,
  setMeasureProvider,
} from "../../layout-engine/measure/measureProvider";

const fixtureUrl = (name: string) =>
  new URL(`../../docx/__tests__/__fixtures__/corpus/${name}`, import.meta.url);

const FIXED_ADVANCE_RATIO = 0.5;

const installFixedWidthProvider = (): void => {
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
};

const listFor = async (name: string): Promise<DisplayList> => {
  const bytes = await Bun.file(fixtureUrl(name)).arrayBuffer();
  const laidOut = await layoutDocxHeadless(bytes, { pageGap: 0 });
  if (laidOut.isErr()) {
    throw laidOut.error;
  }
  const { layout, blockLookup, documentFeatures, furniture, embeddedFonts } = laidOut.value;
  return buildDisplayList({
    layout,
    blockLookup,
    documentFeatures,
    embeddedFonts,
    ...furniture,
  });
};

const flatten = (primitives: readonly DisplayPrimitive[]): readonly DisplayPrimitive[] =>
  primitives.flatMap((primitive) =>
    primitive.kind === "clipGroup" ||
    primitive.kind === "rotateGroup" ||
    primitive.kind === "opacityGroup"
      ? [primitive, ...flatten(primitive.children)]
      : [primitive],
  );

const runsOn = (list: DisplayList, pageIndex: number): readonly DisplayGlyphRun[] =>
  flatten(list.pages[pageIndex]?.primitives ?? []).flatMap((primitive) =>
    primitive.kind === "glyphRun" ? [primitive] : [],
  );

const runContaining = (
  runs: readonly DisplayGlyphRun[],
  text: string,
): DisplayGlyphRun | undefined => runs.find((run) => run.text.includes(text));

describe("a run's model range names the story it belongs to", () => {
  const installed = getMeasureProvider();

  afterEach(() => {
    setMeasureProvider(installed);
  });

  test("a header run addresses the header part the page selected", async () => {
    installFixedWidthProvider();
    const runs = runsOn(await listFor("step3-header-footer-fields.docx"), 1);

    const header = runContaining(runs, "Share transfer agreement");
    expect(header?.pmRange).toBeDefined();
    expect(header?.pmRange?.story).toEqual({ kind: "header", rId: "rId11" });
    expect(header?.pmRange?.end).toBeGreaterThan(header?.pmRange?.start ?? 0);

    // The first page selects the other part, and says so.
    const firstPage = runContaining(
      runsOn(await listFor("step3-header-footer-fields.docx"), 0),
      "Confidential",
    );
    expect(firstPage?.pmRange?.story).toEqual({ kind: "header", rId: "rId13" });
  });

  test("a substituted field value carries no range, in any story", async () => {
    installFixedWidthProvider();
    const runs = runsOn(await listFor("step3-header-footer-fields.docx"), 1);

    // The footer of this document is one `PAGE` field and nothing else. Its
    // digits are the layout's answer rather than characters the document has,
    // so there is no position to put a caret at inside them, and a range would
    // address whatever happened to sit at those numbers.
    const pageNumber = runs.filter((run) => run.text === "2");
    expect(pageNumber).not.toHaveLength(0);
    expect(pageNumber.every((run) => run.pmRange === undefined)).toBe(true);

    // The literal text around the header's fields is addressable, so a run
    // going unaddressed is a property of the field, not of the story.
    expect(runContaining(runs, "Share transfer agreement")?.pmRange).toBeDefined();
  });

  test("a footnote body addresses its own note, by id", async () => {
    installFixedWidthProvider();
    const runs = runsOn(await listFor("step3-footnotes.docx"), 0);

    const note = runContaining(runs, "Counted from the day");
    expect(note?.pmRange?.story).toEqual({ kind: "footnote", id: 2 });
    const second = runContaining(runs, "An invoice meeting");
    expect(second?.pmRange?.story).toEqual({ kind: "footnote", id: 3 });
  });

  test("body runs on the same page address the body", async () => {
    installFixedWidthProvider();
    const runs = runsOn(await listFor("step3-footnotes.docx"), 0);

    const body = runContaining(runs, "The purchase price");
    expect(body?.pmRange?.story).toEqual({ kind: "body" });
  });

  test("no two stories on a page hand out the same range", async () => {
    installFixedWidthProvider();
    const runs = runsOn(await listFor("step3-footnotes.docx"), 0);

    // A body run and a note run can carry the same numbers; what must never
    // repeat is the pair, or a surface could not tell which document a click
    // belongs to.
    const keys = runs.flatMap((run) =>
      run.pmRange === undefined
        ? []
        : [`${JSON.stringify(run.pmRange.story)}:${String(run.pmRange.start)}`],
    );
    expect(new Set(keys).size).toBe(keys.length);
  });
});
