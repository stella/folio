import { describe, expect, test } from "bun:test";

import { compareGeoms, mergeVisualRows } from "../compare";
import { DEFAULT_TOLERANCES } from "../config";
import { normalizeLineText } from "../textNorm";
import type { DocGeom, LineBox, PageGeom } from "../types";

const makeLine = (overrides: Partial<LineBox> & Pick<LineBox, "text">): LineBox => ({
  normText: normalizeLineText(overrides.text),
  xPt: 72,
  yPt: 72,
  widthPt: 100,
  heightPt: 12,
  region: "body",
  ...overrides,
});

const makePage = (overrides: Partial<PageGeom> & Pick<PageGeom, "lines">): PageGeom => ({
  number: 1,
  widthPt: 612,
  heightPt: 792,
  ...overrides,
});

const makeDoc = (source: "word" | "folio", pages: PageGeom[]): DocGeom => ({
  source,
  file: "/test.docx",
  pages,
  meta: {},
});

describe("compareGeoms", () => {
  test("merges tabbed legal markers on the same visual row", () => {
    const merged = mergeVisualRows([
      makeLine({ text: "(b)", xPt: 90, yPt: 100, widthPt: 12.8, heightPt: 12 }),
      makeLine({
        text: "Liquidity Event. If there is a Liquidity Event",
        xPt: 126,
        yPt: 100,
        widthPt: 250,
        heightPt: 12,
      }),
    ]);

    expect(merged).toHaveLength(1);
    expect(merged[0]?.normText).toBe(
      normalizeLineText("(b) Liquidity Event. If there is a Liquidity Event"),
    );
  });

  test("keeps wider same-row columns separate", () => {
    const merged = mergeVisualRows([
      makeLine({ text: "Left cell", xPt: 90, yPt: 100, widthPt: 40, heightPt: 12 }),
      makeLine({ text: "Right cell", xPt: 155, yPt: 100, widthPt: 60, heightPt: 12 }),
    ]);

    expect(merged).toHaveLength(2);
  });

  test("merges wide tab segments from the same logical line", () => {
    const merged = mergeVisualRows([
      makeLine({
        text: "Left segment",
        xPt: 45,
        yPt: 100,
        widthPt: 90,
        logicalLineGroup: "layout-line:7",
      }),
      makeLine({
        text: "Right segment",
        xPt: 350,
        yPt: 100,
        widthPt: 100,
        logicalLineGroup: "layout-line:7",
      }),
    ]);

    expect(merged).toHaveLength(1);
    expect(merged[0]?.normText).toBe(normalizeLineText("Left segment Right segment"));
    expect(merged[0]?.logicalLineGroup).toBe("layout-line:7");
  });

  test("keeps wide same-row segments from different logical lines separate", () => {
    const merged = mergeVisualRows([
      makeLine({
        text: "Left column",
        xPt: 45,
        yPt: 100,
        widthPt: 90,
        logicalLineGroup: "layout-line:7",
      }),
      makeLine({
        text: "Right column",
        xPt: 350,
        yPt: 100,
        widthPt: 100,
        logicalLineGroup: "layout-line:8",
      }),
    ]);

    expect(merged).toHaveLength(2);
  });

  test("merges nearby text from different table cells like reference PDF boxes", () => {
    const merged = mergeVisualRows([
      makeLine({
        text: "Left cell",
        xPt: 90,
        yPt: 100,
        widthPt: 60,
        heightPt: 12,
        visualGroup: "table-cell:0",
      }),
      makeLine({
        text: "Right cell",
        xPt: 155,
        yPt: 100,
        widthPt: 60,
        heightPt: 12,
        visualGroup: "table-cell:1",
      }),
    ]);

    expect(merged).toHaveLength(1);
    expect(merged[0]?.normText).toBe(normalizeLineText("Left cell Right cell"));
  });

  test("merges nearby table and ungrouped text without Folio-only asymmetry", () => {
    const merged = mergeVisualRows([
      makeLine({ text: "Body text", xPt: 90, yPt: 100, widthPt: 60, heightPt: 12 }),
      makeLine({
        text: "Cell text",
        xPt: 155,
        yPt: 100,
        widthPt: 60,
        heightPt: 12,
        visualGroup: "table-cell:0",
      }),
    ]);

    expect(merged).toHaveLength(1);
  });

  test("keeps nearby boxes from different semantic regions separate", () => {
    const merged = mergeVisualRows([
      makeLine({ text: "Header", region: "header", xPt: 90, yPt: 100, widthPt: 60 }),
      makeLine({ text: "Body", region: "body", xPt: 155, yPt: 100, widthPt: 60 }),
    ]);

    expect(merged).toHaveLength(2);
  });

  test("normalizes adjacent table cells symmetrically across extractors", () => {
    const word = makeDoc("word", [
      makePage({
        lines: [
          makeLine({ text: "počet", xPt: 90, yPt: 100, widthPt: 30, heightPt: 12 }),
          makeLine({ text: "celkem", xPt: 125, yPt: 100, widthPt: 35, heightPt: 12 }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({
            text: "počet",
            xPt: 90,
            yPt: 100,
            widthPt: 30,
            heightPt: 12,
            visualGroup: "table-cell:0",
          }),
          makeLine({
            text: "celkem",
            xPt: 125,
            yPt: 100,
            widthPt: 35,
            heightPt: 12,
            visualGroup: "table-cell:1",
          }),
        ],
      }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.divergences.filter((divergence) => divergence.kind === "line-break")).toEqual([]);
    expect(result.score).toBe(1);
  });

  test("merges standalone list markers with their body across marker-sized gaps", () => {
    const merged = mergeVisualRows([
      makeLine({ text: "(i)", xPt: 108, yPt: 100, widthPt: 10.4, heightPt: 12 }),
      makeLine({
        text: "Junior to payment",
        xPt: 144,
        yPt: 100,
        widthPt: 90,
        heightPt: 12,
      }),
    ]);

    expect(merged).toHaveLength(1);
    expect(merged[0]?.normText).toBe(normalizeLineText("(i) Junior to payment"));
  });

  test("identical docs score 1 with zero divergences and zero median offset", () => {
    const pages = [
      makePage({
        lines: [
          makeLine({ text: "Hello world", yPt: 72 }),
          makeLine({ text: "Second line", yPt: 90 }),
        ],
      }),
    ];
    const word = makeDoc("word", pages);
    const folio = makeDoc(
      "folio",
      pages.map((page) =>
        Object.assign({}, page, {
          lines: page.lines.map((line) => Object.assign({}, line)),
        }),
      ),
    );

    const result = compareGeoms(word, folio);

    expect(result.score).toBe(1);
    expect(result.divergences).toEqual([]);
    expect(result.medianYOffsetPt).toBe(0);
    expect(result.totalWordLines).toBe(2);
    expect(result.matchedLines).toBe(2);
  });

  test("constant +6pt y offset on every folio line is absorbed as an extractor baseline", () => {
    const word = makeDoc("word", [
      makePage({
        lines: [
          makeLine({ text: "Hello world", yPt: 72 }),
          makeLine({ text: "Second line", yPt: 90 }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({ text: "Hello world", yPt: 78 }),
          makeLine({ text: "Second line", yPt: 96 }),
        ],
      }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.divergences.filter((d) => d.kind === "y-drift")).toEqual([]);
    expect(result.score).toBe(1);
    expect(result.medianYOffsetPt).toBeCloseTo(6);
  });

  test("a single outlier offset (rest at 0) is reported as y-drift on just that line", () => {
    const word = makeDoc("word", [
      makePage({
        lines: [
          makeLine({ text: "Alpha line", yPt: 72 }),
          makeLine({ text: "Bravo line", yPt: 90 }),
          makeLine({ text: "Charlie line", yPt: 108 }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({ text: "Alpha line", yPt: 72 }),
          makeLine({ text: "Bravo line", yPt: 96 }), // +6pt outlier
          makeLine({ text: "Charlie line", yPt: 108 }),
        ],
      }),
    ]);

    const result = compareGeoms(word, folio);

    const yDrifts = result.divergences.filter((d) => d.kind === "y-drift");
    expect(yDrifts).toHaveLength(1);
    expect(yDrifts[0]).toMatchObject({ kind: "y-drift", text: "Bravo line" });
    expect(result.medianYOffsetPt).toBe(0);
    expect(result.score).toBeCloseTo(2 / 3);
  });

  test("an outlier on the first line is attributed to the first line", () => {
    const texts = ["Alpha line", "Bravo line", "Charlie line"];
    const word = makeDoc("word", [
      makePage({
        lines: texts.map((text, index) => makeLine({ text, yPt: 72 + index * 18 })),
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: texts.map((text, index) =>
          makeLine({ text, yPt: 72 + index * 18 + (index === 0 ? 6 : 0) }),
        ),
      }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.divergences.filter((item) => item.kind === "y-drift")).toEqual([
      { kind: "y-drift", page: 1, text: "Alpha line", residualPt: 6 },
    ]);
    expect(result.score).toBeCloseTo(2 / 3);
  });

  test("persistent offset transitions stay localized when the page median shifts", () => {
    const texts = ["Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot"];
    const word = makeDoc("word", [
      makePage({
        lines: texts.map((text, index) => makeLine({ text, yPt: 72 + index * 18 })),
      }),
    ]);
    const makeFolio = (offsets: number[]) =>
      makeDoc("folio", [
        makePage({
          lines: texts.map((text, index) =>
            makeLine({ text, yPt: 72 + index * 18 + (offsets[index] ?? 0) }),
          ),
        }),
      ]);

    const before = compareGeoms(word, makeFolio([-3, -3, 1, 1, -4, -4]));
    const after = compareGeoms(word, makeFolio([-3, -3, 1, 1, 1, 1]));
    const beforeYDrifts = before.divergences.filter((item) => item.kind === "y-drift");
    const afterYDrifts = after.divergences.filter((item) => item.kind === "y-drift");

    expect(beforeYDrifts).toEqual([
      { kind: "y-drift", page: 1, text: "Charlie", residualPt: 4 },
      { kind: "y-drift", page: 1, text: "Echo", residualPt: -5 },
    ]);
    expect(afterYDrifts).toEqual([{ kind: "y-drift", page: 1, text: "Charlie", residualPt: 4 }]);
    expect(after.score).toBeGreaterThan(before.score);
    expect(after.medianYOffsetPt).not.toBe(before.medianYOffsetPt);
  });

  test("punctuation-only rows do not create y-drift or bias the extractor offset", () => {
    const word = makeDoc("word", [
      makePage({
        lines: [
          makeLine({ text: "____________________", yPt: 72 }),
          makeLine({ text: "....................", yPt: 90 }),
          makeLine({ text: "Baseline text", yPt: 108 }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({ text: "____________________", yPt: 82 }),
          makeLine({ text: "....................", yPt: 100 }),
          makeLine({ text: "Baseline text", yPt: 114 }),
        ],
      }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.divergences.filter((divergence) => divergence.kind === "y-drift")).toEqual([]);
    expect(result.medianYOffsetPt).toBe(6);
    expect(result.score).toBe(1);
  });

  test("punctuation-only rows still report horizontal geometry drift", () => {
    const word = makeDoc("word", [
      makePage({
        lines: [
          makeLine({ text: "Baseline text", yPt: 72 }),
          makeLine({ text: "____________________", yPt: 90, widthPt: 100 }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({ text: "Baseline text", yPt: 72 }),
          makeLine({ text: "____________________", yPt: 96, widthPt: 120 }),
        ],
      }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.divergences.filter((divergence) => divergence.kind === "y-drift")).toEqual([]);
    expect(
      result.divergences.filter((divergence) => divergence.kind === "width-drift"),
    ).toHaveLength(1);
    expect(result.score).toBe(0.5);
  });

  test("absorbs each page's extractor offset independently", () => {
    const word = makeDoc("word", [
      makePage({
        number: 1,
        lines: [
          makeLine({ text: "Page one A", yPt: 72 }),
          makeLine({ text: "Page one B", yPt: 90 }),
        ],
      }),
      makePage({
        number: 2,
        lines: [
          makeLine({ text: "Page two A", yPt: 72 }),
          makeLine({ text: "Page two B", yPt: 90 }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        number: 1,
        lines: [
          makeLine({ text: "Page one A", yPt: 69 }),
          makeLine({ text: "Page one B", yPt: 87 }),
        ],
      }),
      makePage({
        number: 2,
        lines: [
          makeLine({ text: "Page two A", yPt: 84 }),
          makeLine({ text: "Page two B", yPt: 102 }),
        ],
      }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.divergences.filter((divergence) => divergence.kind === "y-drift")).toEqual([]);
    expect(result.score).toBe(1);
    expect(result.medianYOffsetPt).toBeCloseTo(4.5);
  });

  test("uses Folio regions to absorb header and body extractor offsets independently", () => {
    const word = makeDoc("word", [
      makePage({
        lines: [
          makeLine({ text: "Header", yPt: 20 }),
          makeLine({ text: "Body A", yPt: 72 }),
          makeLine({ text: "Body B", yPt: 90 }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({ text: "Header", region: "header", yPt: 17 }),
          makeLine({ text: "Body A", region: "body", yPt: 84 }),
          makeLine({ text: "Body B", region: "body", yPt: 102 }),
        ],
      }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.divergences.filter((divergence) => divergence.kind === "y-drift")).toEqual([]);
    expect(result.score).toBe(1);
    expect(result.medianYOffsetPt).toBeCloseTo(4.5);
  });

  test("a reference line split into two folio lines reconciles as a single line-break", () => {
    const word = makeDoc("word", [
      makePage({ lines: [makeLine({ text: "The quick brown fox jumps" })] }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({ text: "The quick brown", yPt: 72 }),
          makeLine({ text: "fox jumps", yPt: 90 }),
        ],
      }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.divergences).toHaveLength(1);
    expect(result.divergences[0]).toMatchObject({
      kind: "line-break",
      wordTexts: ["The quick brown fox jumps"],
      folioTexts: ["The quick brown", "fox jumps"],
    });
    expect(
      result.divergences.some((d) => d.kind === "missing-line" || d.kind === "extra-line"),
    ).toBe(false);
  });

  test("accepts equivalent same-row segmentation when union geometry agrees", () => {
    const word = makeDoc("word", [
      makePage({
        lines: [
          makeLine({ text: "Left one", xPt: 72, yPt: 72, widthPt: 60 }),
          makeLine({ text: "Right one", xPt: 300, yPt: 72, widthPt: 60 }),
          makeLine({ text: "Left two", xPt: 72, yPt: 100, widthPt: 60 }),
          makeLine({ text: "Right two", xPt: 300, yPt: 100, widthPt: 60 }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({ text: "Left one Right one", xPt: 72, yPt: 78, widthPt: 288 }),
          makeLine({ text: "Left two Right two", xPt: 72, yPt: 106, widthPt: 288 }),
        ],
      }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.divergences).toEqual([]);
    expect(result.matchedLines).toBe(4);
    expect(result.medianYOffsetPt).toBe(6);
    expect(result.score).toBe(1);
  });

  test("reconciles one segmented visual row inside a larger alignment gap", () => {
    const reference = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({ text: "Short", xPt: 72, yPt: 72, widthPt: 30, region: "unknown" }),
          makeLine({
            text: "Right segment",
            xPt: 180,
            yPt: 72,
            widthPt: 90,
            region: "unknown",
          }),
          makeLine({
            text: "Unmatched row",
            xPt: 72,
            yPt: 100,
            widthPt: 80,
            region: "unknown",
          }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [makeLine({ text: "Short Right segment", xPt: 72, yPt: 78, widthPt: 198 })],
      }),
    ]);

    const result = compareGeoms(reference, folio);

    expect(result.divergences).toEqual([{ kind: "missing-line", page: 1, text: "Unmatched row" }]);
    expect(result.matchedLines).toBe(2);
    expect(result.medianYOffsetPt).toBe(6);
    expect(result.score).toBeCloseTo(2 / 3);
  });

  test("a normalized 1-to-1 gap is treated as a match, not a line-break", () => {
    const word = makeDoc("word", [
      makePage({
        lines: [
          makeLine({ text: "1.", xPt: 10 }),
          makeLine({ text: "Definitions ................ 2", xPt: 40 }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [makeLine({ text: "1.", xPt: 10 }), makeLine({ text: "Definitions … 2", xPt: 40 })],
      }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.divergences).toEqual([]);
    expect(result.score).toBe(1);
  });

  test("an extra folio page with a shifted line reports page-count and pagination", () => {
    const word = makeDoc("word", [
      makePage({
        number: 1,
        lines: [makeLine({ text: "Line one", yPt: 72 }), makeLine({ text: "Line two", yPt: 90 })],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({ number: 1, lines: [makeLine({ text: "Line one", yPt: 72 })] }),
      makePage({ number: 2, lines: [makeLine({ text: "Line two", yPt: 72 })] }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.divergences[0]).toEqual({ kind: "page-count", word: 1, folio: 2 });
    const pagination = result.divergences.find((d) => d.kind === "pagination");
    expect(pagination).toMatchObject({
      kind: "pagination",
      text: "Line two",
      wordPage: 1,
      folioPage: 2,
    });
  });

  test("a reference-only line is missing; a folio-only line is extra", () => {
    const word = makeDoc("word", [
      makePage({
        lines: [
          makeLine({ text: "Alpha", yPt: 72 }),
          makeLine({ text: "Zephyr quotient marginalia", yPt: 90 }),
          makeLine({ text: "Bravo", yPt: 108 }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({ text: "Alpha", yPt: 72 }),
          makeLine({ text: "Umbrella coefficient tapestry", yPt: 90 }),
          makeLine({ text: "Bravo", yPt: 108 }),
        ],
      }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.divergences).toContainEqual({
      kind: "missing-line",
      page: 1,
      text: "Zephyr quotient marginalia",
    });
    expect(result.divergences).toContainEqual({
      kind: "extra-line",
      page: 1,
      text: "Umbrella coefficient tapestry",
    });
  });

  test("x-drift: flagged beyond tolerance, not flagged within tolerance", () => {
    const word = makeDoc("word", [
      makePage({
        lines: [
          makeLine({ text: "Drift beyond tolerance", xPt: 72, yPt: 72 }),
          makeLine({ text: "Drift within tolerance", xPt: 72, yPt: 90 }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({
            text: "Drift beyond tolerance",
            xPt: 72 + DEFAULT_TOLERANCES.xPt + 0.5,
            yPt: 72,
          }),
          makeLine({
            text: "Drift within tolerance",
            xPt: 72 + DEFAULT_TOLERANCES.xPt - 0.5,
            yPt: 90,
          }),
        ],
      }),
    ]);

    const result = compareGeoms(word, folio);

    const xDrifts = result.divergences.filter((d) => d.kind === "x-drift");
    expect(xDrifts).toHaveLength(1);
    expect(xDrifts[0]).toMatchObject({ kind: "x-drift", text: "Drift beyond tolerance" });
  });

  test("width-drift: wide line passes on relative tolerance, narrow line fails on absolute", () => {
    const wideWordWidth = 400;
    const wideDelta = DEFAULT_TOLERANCES.widthPt + 0.5; // fails absolute, passes relative on a wide line
    const narrowWordWidth = 50;
    const narrowDelta = DEFAULT_TOLERANCES.widthPt + 0.5; // fails both absolute and relative on a narrow line
    expect(wideDelta).toBeLessThanOrEqual(DEFAULT_TOLERANCES.widthRelative * wideWordWidth);
    expect(narrowDelta).toBeGreaterThan(DEFAULT_TOLERANCES.widthRelative * narrowWordWidth);

    const word = makeDoc("word", [
      makePage({
        lines: [
          makeLine({
            text: "Wide line passes on relative tolerance",
            widthPt: wideWordWidth,
            yPt: 72,
          }),
          makeLine({ text: "Narrow line fails", widthPt: narrowWordWidth, yPt: 90 }),
        ],
      }),
    ]);
    const folio = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({
            text: "Wide line passes on relative tolerance",
            widthPt: wideWordWidth + wideDelta,
            yPt: 72,
          }),
          makeLine({ text: "Narrow line fails", widthPt: narrowWordWidth + narrowDelta, yPt: 90 }),
        ],
      }),
    ]);

    const result = compareGeoms(word, folio);

    const widthDrifts = result.divergences.filter((d) => d.kind === "width-drift");
    expect(widthDrifts).toHaveLength(1);
    expect(widthDrifts[0]).toMatchObject({ kind: "width-drift", text: "Narrow line fails" });
  });

  test("near-identical (but not equal) text on a matched pair is a text-mismatch", () => {
    const word = makeDoc("word", [
      makePage({ lines: [makeLine({ text: "The quick brown fox" })] }),
    ]);
    const folio = makeDoc("folio", [
      makePage({ lines: [makeLine({ text: "The quick brown fox." })] }),
    ]);

    const result = compareGeoms(word, folio);

    expect(result.matchedLines).toBe(1);
    expect(result.divergences).toContainEqual({
      kind: "text-mismatch",
      page: 1,
      wordText: "The quick brown fox",
      folioText: "The quick brown fox.",
    });
  });

  test("empty vs empty scores 1; empty word vs non-empty folio scores 0 with extra-lines", () => {
    const emptyWord = makeDoc("word", [makePage({ lines: [] })]);
    const emptyFolio = makeDoc("folio", [makePage({ lines: [] })]);
    const emptyResult = compareGeoms(emptyWord, emptyFolio);

    expect(emptyResult.score).toBe(1);
    expect(emptyResult.divergences).toEqual([]);
    expect(emptyResult.totalWordLines).toBe(0);

    const folioOnly = makeDoc("folio", [
      makePage({
        lines: [
          makeLine({ text: "Only in folio one", yPt: 100 }),
          makeLine({ text: "Only in folio two", yPt: 120 }),
        ],
      }),
    ]);
    const mixedResult = compareGeoms(emptyWord, folioOnly);

    expect(mixedResult.score).toBe(0);
    expect(mixedResult.divergences).toHaveLength(2);
    expect(mixedResult.divergences.every((d) => d.kind === "extra-line")).toBe(true);
  });
});
