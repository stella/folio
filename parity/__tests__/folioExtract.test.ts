import { describe, expect, test } from "bun:test";

import { PX_TO_PT } from "../config";
import {
  CLEAN_SCREENSHOT_CSS,
  computeZoomFactor,
  formatServerStartFailure,
  meaningfulTextRange,
  parseCssFontFamilies,
  parseFirstFontFamily,
  toPageGeom,
} from "../folioExtract";
import type { RawLine, RawPage } from "../folioExtract";

const rect = (left: number, top: number, width: number, height: number) => ({
  left,
  top,
  width,
  height,
});

describe("clean screenshot style", () => {
  test("hides editor and playground chrome that can overlap a page", () => {
    expect(CLEAN_SCREENSHOT_CSS).toContain('[data-folio-toolbar="true"]');
    expect(CLEAN_SCREENSHOT_CSS).toContain('[data-testid="playground-controls"]');
  });

  test("keeps the page and its document content visible", () => {
    expect(CLEAN_SCREENSHOT_CSS).toContain(".layout-page,");
    expect(CLEAN_SCREENSHOT_CSS).toContain(".layout-page *");
  });
});

describe("playground startup diagnostics", () => {
  test("includes an early exit code and captured output", () => {
    expect(formatServerStartFailure("http://localhost:4200", 90_000, 127, "vite: not found")).toBe(
      "playground dev server exited with code 127 before becoming ready at http://localhost:4200\n" +
        "Playground output:\n" +
        "vite: not found",
    );
  });

  test("reports a quiet startup timeout without an empty output section", () => {
    expect(formatServerStartFailure("http://localhost:4200", 90_000, undefined, "")).toBe(
      "playground dev server did not become ready at http://localhost:4200 within 90000ms",
    );
  });
});

describe("parseCssFontFamilies", () => {
  test("preserves commas inside quoted family names", () => {
    expect(parseCssFontFamilies('"Definitely, Missing Font", Arial, sans-serif')).toEqual([
      "Definitely, Missing Font",
      "Arial",
      "sans-serif",
    ]);
  });

  test("handles single-quoted and unquoted families", () => {
    expect(parseCssFontFamilies("'Times New Roman', serif")).toEqual(["Times New Roman", "serif"]);
  });
});

const makeRawLine = (overrides: Partial<RawLine> & Pick<RawLine, "text" | "rect">): RawLine => ({
  region: "body",
  ...overrides,
});

const makeRawPage = (overrides: Partial<RawPage> & Pick<RawPage, "lines">): RawPage => ({
  pageNumber: 1,
  domIndex: 0,
  pageRect: rect(0, 0, 816, 1056),
  offsetWidth: 816,
  offsetHeight: 1056,
  ...overrides,
});

describe("meaningfulTextRange", () => {
  test("excludes surrounding whitespace and invisible controls from ink bounds", () => {
    expect(meaningfulTextRange("   \u00ad\u200bč. NPU-450/74723/2025 \t")).toEqual({
      start: 5,
      end: 26,
    });
    expect(meaningfulTextRange(" \t\u00ad\u200b")).toBeNull();
  });
});

describe("computeZoomFactor", () => {
  test("is 1 when offsetWidth matches the visual rect width (no CSS zoom)", () => {
    expect(computeZoomFactor(816, 816)).toBe(1);
  });

  test("reflects a CSS transform zoom (visual rect shrinks relative to layout px)", () => {
    // Page is laid out at 816 layout px but rendered at 50% zoom, so the
    // visual bounding rect is 408px wide; zoomFactor normalizes that back.
    expect(computeZoomFactor(816, 408)).toBe(2);
  });

  test("guards divide-by-zero for a zero-width page rect", () => {
    expect(computeZoomFactor(816, 0)).toBe(1);
  });

  test("guards a negative or non-finite page rect width", () => {
    expect(computeZoomFactor(816, -5)).toBe(1);
    expect(computeZoomFactor(816, Number.NaN)).toBe(1);
  });
});

describe("parseFirstFontFamily", () => {
  test("takes the first family and strips double quotes", () => {
    expect(parseFirstFontFamily('"Calibri", "Arial", sans-serif')).toBe("Calibri");
  });

  test("strips single quotes", () => {
    expect(parseFirstFontFamily("'Times New Roman', serif")).toBe("Times New Roman");
  });

  test("passes through an unquoted single family unchanged", () => {
    expect(parseFirstFontFamily("Arial")).toBe("Arial");
  });

  test("returns undefined for missing or empty input", () => {
    expect(parseFirstFontFamily(undefined)).toBeUndefined();
    expect(parseFirstFontFamily("")).toBeUndefined();
  });
});

describe("toPageGeom", () => {
  test("converts px geometry to pt at zoomFactor 1 (no CSS zoom)", () => {
    // PX_TO_PT = 0.75: a page laid out at 816x1056 px is 612x792pt (US Letter).
    const rawPage = makeRawPage({
      lines: [makeRawLine({ text: "Hello", rect: rect(96, 96, 200, 20) })],
    });

    const page = toPageGeom(rawPage);

    expect(page.widthPt).toBeCloseTo(612, 5);
    expect(page.heightPt).toBeCloseTo(792, 5);
    expect(page.lines).toHaveLength(1);
    expect(page.lines[0]?.xPt).toBeCloseTo(96 * PX_TO_PT, 5);
    expect(page.lines[0]?.yPt).toBeCloseTo(96 * PX_TO_PT, 5);
    expect(page.lines[0]?.widthPt).toBeCloseTo(200 * PX_TO_PT, 5);
    expect(page.lines[0]?.heightPt).toBeCloseTo(20 * PX_TO_PT, 5);
  });

  test("normalizes coordinates against a CSS-zoomed page (zoomFactor != 1)", () => {
    // Page laid out at 816 layout px, rendered at 50% zoom (visual rect
    // 408px wide). A line whose visual left edge is 48px from the visual
    // page left should resolve to the same pt offset as if unzoomed: 48px
    // visual * zoomFactor(2) * PX_TO_PT = 72pt, matching a 96px layout
    // offset at zoomFactor 1 (96 * 0.75 = 72pt).
    const rawPage = makeRawPage({
      pageRect: rect(100, 50, 408, 528),
      offsetWidth: 816,
      offsetHeight: 1056,
      lines: [makeRawLine({ text: "Zoomed", rect: rect(148, 74, 100, 10) })],
    });

    const page = toPageGeom(rawPage);

    expect(page.widthPt).toBeCloseTo(612, 5);
    expect(page.lines[0]?.xPt).toBeCloseTo(72, 5);
    expect(page.lines[0]?.yPt).toBeCloseTo(36, 5);
    expect(page.lines[0]?.widthPt).toBeCloseTo(150, 5);
  });

  test("drops zero-size lines", () => {
    const rawPage = makeRawPage({
      lines: [
        makeRawLine({ text: "Visible", rect: rect(0, 0, 100, 10) }),
        makeRawLine({ text: "Zero width", rect: rect(0, 0, 0, 10) }),
        makeRawLine({ text: "Zero height", rect: rect(0, 0, 100, 0) }),
      ],
    });

    const page = toPageGeom(rawPage);

    expect(page.lines).toHaveLength(1);
    expect(page.lines[0]?.text).toBe("Visible");
  });

  test("drops lines whose normalized text is empty (whitespace-only / soft hyphen only)", () => {
    const rawPage = makeRawPage({
      lines: [
        makeRawLine({ text: "   ", rect: rect(0, 0, 100, 10) }),
        makeRawLine({ text: "­​", rect: rect(0, 20, 100, 10) }),
        makeRawLine({ text: "Kept", rect: rect(0, 40, 100, 10) }),
      ],
    });

    const page = toPageGeom(rawPage);

    expect(page.lines).toHaveLength(1);
    expect(page.lines[0]?.normText).toBe("Kept");
  });

  test("sorts lines by (yPt, xPt)", () => {
    const rawPage = makeRawPage({
      lines: [
        makeRawLine({ text: "second-line-right", rect: rect(50, 20, 40, 10) }),
        makeRawLine({ text: "first-line", rect: rect(0, 0, 40, 10) }),
        makeRawLine({ text: "second-line-left", rect: rect(0, 20, 40, 10) }),
      ],
    });

    const page = toPageGeom(rawPage);

    expect(page.lines.map((line) => line.text)).toEqual([
      "first-line",
      "second-line-left",
      "second-line-right",
    ]);
  });

  test("maps region straight through from the raw line", () => {
    const rawPage = makeRawPage({
      lines: [
        makeRawLine({ text: "Header text", rect: rect(0, 0, 40, 10), region: "header" }),
        makeRawLine({ text: "Body text", rect: rect(0, 20, 40, 10), region: "body" }),
        makeRawLine({ text: "Footer text", rect: rect(0, 40, 40, 10), region: "footer" }),
      ],
    });

    const page = toPageGeom(rawPage);

    expect(page.lines.map((line) => line.region)).toEqual(["header", "body", "footer"]);
  });

  test("preserves table-cell visual groups", () => {
    const rawPage = makeRawPage({
      lines: [
        makeRawLine({
          text: "Cell text",
          rect: rect(0, 0, 40, 10),
          visualGroup: "table-cell:3",
        }),
      ],
    });

    expect(toPageGeom(rawPage).lines.at(0)?.visualGroup).toBe("table-cell:3");
  });

  test("computes fontName/fontSizePt from raw px values, omitting when absent", () => {
    const rawPage = makeRawPage({
      lines: [
        makeRawLine({
          text: "Styled",
          rect: rect(0, 0, 40, 10),
          fontFamilyRaw: '"Times New Roman", serif',
          fontSizePx: 16,
        }),
        makeRawLine({ text: "Unstyled", rect: rect(0, 20, 40, 10) }),
      ],
    });

    const page = toPageGeom(rawPage);

    expect(page.lines[0]?.fontName).toBe("Times New Roman");
    expect(page.lines[0]?.fontSizePt).toBeCloseTo(12, 5);
    expect(page.lines[1]?.fontName).toBeUndefined();
    expect(page.lines[1]?.fontSizePt).toBeUndefined();
  });

  test("preserves the raw page number", () => {
    const rawPage = makeRawPage({
      pageNumber: 3,
      lines: [makeRawLine({ text: "Page three", rect: rect(0, 0, 40, 10) })],
    });

    expect(toPageGeom(rawPage).number).toBe(3);
  });
});
