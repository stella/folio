import { describe, expect, test } from "bun:test";

import type { ParagraphBlock, Run } from "../types";
import {
  smallCapsAwareCharWidth,
  withFakeTextMeasure,
  fixedCharWidth,
} from "./__tests__/fakeTextMeasure";
import { hashParagraphBlock } from "./cache";
import { buildFontString } from "./measureHelpers";
import { clampFloatingWrapMargins, getRunCharWidths, measureParagraph } from "./measureParagraph";
import { getFontMetrics, measureTextWidth } from "./measureProvider";

const PT_TO_PX = 96 / 72;

const fakeMeasure = { charWidth: smallCapsAwareCharWidth };

describe("text measurement cache", () => {
  test("reuses canvas text width measurements for identical text and style", () => {
    withFakeTextMeasure((getMeasureCount) => {
      const style = { fontFamily: "Arial", fontSize: 11 };

      expect(measureTextWidth("Repeated legal text", style)).toBe(
        measureTextWidth("Repeated legal text", style),
      );
      expect(getMeasureCount()).toBe(1);
    }, fakeMeasure);
  });

  test("keeps horizontal scale in the text width cache key", () => {
    withFakeTextMeasure((getMeasureCount) => {
      const text = "scaled";
      const normalWidth = measureTextWidth(text, {
        fontFamily: "Arial",
        fontSize: 11,
      });
      const scaledWidth = measureTextWidth(text, {
        fontFamily: "Arial",
        fontSize: 11,
        horizontalScale: 150,
      });

      expect(scaledWidth).toBe(normalWidth * 1.5);
      expect(getMeasureCount()).toBe(2);
    }, fakeMeasure);
  });

  test("keeps enabled kerning in the text width cache key", () => {
    withFakeTextMeasure((getMeasureCount) => {
      const text = "AVAV";
      measureTextWidth(text, { fontFamily: "Arial", fontSize: 11, kerning: false });
      measureTextWidth(text, { fontFamily: "Arial", fontSize: 11, kerning: true });

      expect(getMeasureCount()).toBe(2);
    }, fakeMeasure);
  });
});

describe("run kerning threshold", () => {
  test("disables pair kerning unless the authored threshold is met", () => {
    withFakeTextMeasure(
      () => {
        const runs = [{ kind: "text" as const, text: "AVAV", fontSize: 11 }];
        const charWidth = 4;
        const availableWidth = runs[0]!.text.length * charWidth;
        const base = { kind: "paragraph" as const, id: "kerning", runs };

        expect(measureParagraph(base, availableWidth).lines).toHaveLength(2);
        expect(
          measureParagraph({ ...base, runs: [{ ...runs[0]!, kerningMinPt: 10 }] }, availableWidth)
            .lines,
        ).toHaveLength(1);
        expect(
          measureParagraph({ ...base, runs: [{ ...runs[0]!, kerningMinPt: 12 }] }, availableWidth)
            .lines,
        ).toHaveLength(2);
      },
      {
        charWidth: (_char, _font, fontKerning) => (fontKerning === "normal" ? 4 : 5),
      },
    );
  });
});

describe("font metrics cache", () => {
  test("reuses canvas font metrics for identical font styles", () => {
    withFakeTextMeasure((getMeasureCount) => {
      const style = { fontFamily: "Arial", fontSize: 11 };

      expect(getFontMetrics(style)).toEqual(getFontMetrics(style));
      expect(getMeasureCount()).toBe(1);
    }, fakeMeasure);
  });

  test("keeps font variant in the metrics cache key", () => {
    withFakeTextMeasure((getMeasureCount) => {
      getFontMetrics({ fontFamily: "Arial", fontSize: 11 });
      getFontMetrics({
        fontFamily: "Arial",
        fontSize: 11,
        fontVariant: "small-caps",
      });

      expect(getMeasureCount()).toBe(2);
    }, fakeMeasure);
  });

  test("keeps outline level in the paragraph measurement cache key", () => {
    const paragraph = {
      kind: "paragraph" as const,
      id: "outline-cache",
      runs: [],
    };

    expect(hashParagraphBlock(paragraph)).not.toBe(
      hashParagraphBlock({
        ...paragraph,
        attrs: { outlineLevel: 0, reserveEmptyOutlineHeight: true },
      }),
    );
  });

  test("keeps automatic hyphenation policy in the paragraph measurement cache key", () => {
    const paragraph: ParagraphBlock = {
      kind: "paragraph",
      id: "hyphenation-cache",
      runs: [{ kind: "text", text: "hyphenation", language: { val: "en-US" } }],
    };

    expect(hashParagraphBlock(paragraph)).not.toBe(
      hashParagraphBlock({
        ...paragraph,
        attrs: { automaticHyphenation: { enabled: true } },
      }),
    );
    expect(
      hashParagraphBlock({
        ...paragraph,
        attrs: { automaticHyphenation: { enabled: true } },
      }),
    ).not.toBe(
      hashParagraphBlock({
        ...paragraph,
        attrs: {
          automaticHyphenation: { enabled: true },
          suppressAutoHyphens: true,
        },
      }),
    );
    expect(
      hashParagraphBlock({
        ...paragraph,
        attrs: {
          automaticHyphenation: { enabled: true, hyphenationZoneTwips: 360 },
        },
      }),
    ).not.toBe(
      hashParagraphBlock({
        ...paragraph,
        attrs: {
          automaticHyphenation: { enabled: true, hyphenationZoneTwips: 720 },
        },
      }),
    );
  });
});

describe("empty paragraph line-height floor", () => {
  test("empty paragraph with line=1.0 auto is floored to 1.15 times fontSize", () => {
    const measure = measureParagraph(
      {
        kind: "paragraph",
        id: "t1",
        pmStart: 0,
        pmEnd: 0,
        runs: [],
        attrs: {
          defaultFontSize: 11,
          defaultFontFamily: "Arial Narrow",
          spacing: { line: 1, lineUnit: "multiplier", lineRule: "auto" },
        },
      },
      600,
    );

    expect(measure.totalHeight).toBeCloseTo(11 * PT_TO_PX * 1.15, 1);
  });

  test("empty paragraph with lineRule=exact is not floored", () => {
    const measure = measureParagraph(
      {
        kind: "paragraph",
        id: "t2",
        pmStart: 0,
        pmEnd: 0,
        runs: [],
        attrs: {
          defaultFontSize: 11,
          defaultFontFamily: "Arial Narrow",
          spacing: { line: 8, lineUnit: "px", lineRule: "exact" },
        },
      },
      600,
    );

    expect(measure.totalHeight).toBeCloseTo(8, 1);
  });

  test("empty paragraph includes authored spacing", () => {
    const measure = measureParagraph(
      {
        kind: "paragraph",
        id: "t3",
        pmStart: 0,
        pmEnd: 0,
        runs: [],
        attrs: {
          defaultFontSize: 11,
          defaultFontFamily: "Arial Narrow",
          spacing: {
            before: 5,
            after: 7,
            line: 1,
            lineUnit: "multiplier",
            lineRule: "auto",
          },
        },
      },
      600,
    );

    expect(measure.totalHeight).toBeCloseTo(11 * PT_TO_PX * 1.15 + 12, 1);
  });

  test("empty top-level outline paragraph reserves two line boxes", () => {
    const paragraph = {
      kind: "paragraph" as const,
      id: "t3-outline",
      pmStart: 0,
      pmEnd: 0,
      runs: [],
      attrs: {
        defaultFontSize: 18,
        defaultFontFamily: "Arial",
      },
    };
    const ordinary = measureParagraph(paragraph, 600);
    const outlined = measureParagraph(
      {
        ...paragraph,
        attrs: {
          ...paragraph.attrs,
          outlineLevel: 0,
          reserveEmptyOutlineHeight: true,
        },
      },
      600,
    );

    expect(outlined.lines).toHaveLength(1);
    expect(outlined.lines[0]?.lineHeight).toBeCloseTo((ordinary.lines[0]?.lineHeight ?? 0) * 2, 1);
    expect(outlined.totalHeight).toBeCloseTo(outlined.lines[0]?.lineHeight ?? 0, 1);
  });

  test("outline metadata does not change non-empty paragraph height", () => {
    const base = {
      kind: "paragraph" as const,
      id: "t3-outline-content",
      pmStart: 0,
      pmEnd: 4,
      runs: [{ kind: "text" as const, text: "text" }],
    };

    withFakeTextMeasure(() => {
      const ordinary = measureParagraph(base, 600);
      const outlined = measureParagraph({ ...base, attrs: { outlineLevel: 0 } }, 600);

      expect(outlined.totalHeight).toBe(ordinary.totalHeight);
      expect(outlined.lines).toEqual(ordinary.lines);
    }, fakeMeasure);
  });

  for (const text of ["", " ", "\u00a0"]) {
    test(`visually empty single text run ${JSON.stringify(text)} includes authored spacing`, () => {
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "t3-text-run",
          pmStart: 0,
          pmEnd: 0,
          runs: [{ kind: "text", text }],
          attrs: {
            defaultFontSize: 11,
            defaultFontFamily: "Arial Narrow",
            spacing: {
              before: 5,
              after: 7,
              line: 1,
              lineUnit: "multiplier",
              lineRule: "auto",
            },
          },
        },
        600,
      );

      expect(measure.lines).toHaveLength(1);
      expect(measure.totalHeight).toBeCloseTo((measure.lines[0]?.lineHeight ?? 0) + 12, 1);
    });
  }

  test("suppressed empty paragraph keeps a zero-height anchor", () => {
    const measure = measureParagraph(
      {
        kind: "paragraph",
        id: "t4",
        pmStart: 0,
        pmEnd: 0,
        runs: [],
        attrs: {
          suppressEmptyParagraphHeight: true,
          defaultFontSize: 11,
          defaultFontFamily: "Arial Narrow",
        },
      },
      600,
    );

    expect(measure.totalHeight).toBe(0);
    expect(measure.lines[0]?.lineHeight).toBe(0);
  });

  test("uses paragraph mark metrics for blank hard-break lines", () => {
    withFakeTextMeasure(() => {
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "hard-break-default-metrics",
          runs: [
            { kind: "text", text: "first", fontSize: 9, fontFamily: "Times New Roman" },
            { kind: "lineBreak" },
            { kind: "lineBreak" },
            { kind: "text", text: "last", fontSize: 9, fontFamily: "Times New Roman" },
          ],
          attrs: {
            defaultFontSize: 9,
            defaultFontFamily: "Times New Roman",
            spacing: { line: 1, lineUnit: "multiplier", lineRule: "auto" },
          },
        },
        600,
      );

      expect(measure.lines).toHaveLength(3);
      expect(measure.lines[1]?.lineHeight).toBeCloseTo(measure.lines[0]?.lineHeight ?? 0, 5);
    }, fakeMeasure);
  });
});

describe("measureParagraph cross-run line breaking", () => {
  const style = { fontFamily: "Calibri", fontSize: 11 };
  const width = (text: string): number => measureTextWidth(text, style);
  const paragraph = (runs: Run[]): ParagraphBlock => ({
    kind: "paragraph",
    id: "p1",
    runs,
  });
  const lineStartsAtRun = (
    lines: { fromRun: number; fromChar: number }[],
    runIndex: number,
  ): boolean => lines.some((line) => line.fromRun === runIndex && line.fromChar === 0);
  const lineStartsAt = (
    lines: { fromRun: number; fromChar: number }[],
    runIndex: number,
    charIndex: number,
  ): boolean => lines.some((line) => line.fromRun === runIndex && line.fromChar === charIndex);

  test("keeps an adjacent footnote marker glued to the preceding word", () => {
    withFakeTextMeasure(() => {
      const runs: Run[] = [
        { kind: "text", text: "alpha beta." },
        { kind: "text", text: "1", superscript: true, footnoteRefId: 1 },
        { kind: "text", text: " gamma" },
      ];
      const maxWidth = width("alpha beta.") + width("1") - 0.6;
      const { lines } = measureParagraph(paragraph(runs), maxWidth);

      expect(lineStartsAtRun(lines, 1)).toBe(false);
    }, fakeMeasure);
  });

  test("uses authored no-break spaces across formatted runs", () => {
    withFakeTextMeasure(() => {
      const ordinaryWhitespace: Run[] = [
        { kind: "text", text: "alpha o ", language: { val: "cs-CZ" } },
        { kind: "text", text: "beta", bold: true, language: { val: "cs-CZ" } },
      ];
      const noBreakWhitespace: Run[] = [
        { kind: "text", text: "alpha o\u00A0", language: { val: "cs-CZ" } },
        { kind: "text", text: "beta", bold: true, language: { val: "cs-CZ" } },
      ];
      const ordinaryMeasure = measureParagraph(paragraph(ordinaryWhitespace), width("alpha o"));
      const noBreakMeasure = measureParagraph(paragraph(noBreakWhitespace), width("alpha o"));

      expect(lineStartsAtRun(ordinaryMeasure.lines, 1)).toBe(true);
      expect(lineStartsAt(noBreakMeasure.lines, 0, "alpha ".length)).toBe(true);
      expect(lineStartsAtRun(noBreakMeasure.lines, 1)).toBe(false);
    }, fakeMeasure);
  });

  test("allows a normal wrap when whitespace precedes the footnote marker", () => {
    withFakeTextMeasure(() => {
      const runs: Run[] = [
        { kind: "text", text: "alpha beta " },
        { kind: "text", text: "1", superscript: true, footnoteRefId: 1 },
        { kind: "text", text: " gamma" },
      ];
      const maxWidth = width("alpha beta ") + width("1") - 0.6;
      const { lines } = measureParagraph(paragraph(runs), maxWidth);

      expect(lineStartsAtRun(lines, 1)).toBe(true);
    }, fakeMeasure);
  });

  test("excludes a final collapsible space from paragraph fit width", () => {
    withFakeTextMeasure(() => {
      const text = "parcela č. ";
      const { lines } = measureParagraph(
        paragraph([{ kind: "text", text }]),
        width(text.trimEnd()),
      );

      expect(lines).toHaveLength(1);
      expect(lines[0]?.toChar).toBe(text.length);
      expect(lines[0]?.width).toBe(width(text.trimEnd()));
    }, fakeMeasure);
  });

  test("consumes a break-space after the last visible word on a line", () => {
    withFakeTextMeasure(() => {
      const text = "alpha beta gamma";
      const { lines } = measureParagraph(paragraph([{ kind: "text", text }]), width("alpha beta"));

      expect(lines).toHaveLength(2);
      expect(lines[0]?.toChar).toBe("alpha beta ".length);
      expect(lines[0]?.width).toBe(width("alpha beta"));
    }, fakeMeasure);
  });

  test("does not create a whitespace-only line after hard-breaking a long word", () => {
    withFakeTextMeasure(() => {
      const text = "abcdefgh ";
      const { lines } = measureParagraph(paragraph([{ kind: "text", text }]), width("abcd"));

      expect(lines).toHaveLength(2);
      expect(lines[1]?.toChar).toBe(text.length);
      expect(lines[1]?.width).toBe(width("efgh"));
    }, fakeMeasure);
  });

  test("hard-breaks long words only at grapheme boundaries", () => {
    withFakeTextMeasure(
      () => {
        const family = "👨‍👩‍👧‍👦";
        const text = `${family}x`;
        const { lines } = measureParagraph(paragraph([{ kind: "text", text }]), 10);

        expect(lines).toHaveLength(2);
        expect(lines[0]?.toChar).toBe(family.length);
        expect(lines[1]).toMatchObject({ fromChar: family.length, toChar: text.length });
      },
      { charWidth: fixedCharWidth(5) },
    );
  });

  test("does not soft-wrap overflowed preserved spaces onto their own lines", () => {
    withFakeTextMeasure(() => {
      const spaces = " ".repeat(12);
      const text = `alpha${spaces}beta`;
      const { lines } = measureParagraph(paragraph([{ kind: "text", text }]), width("alpha"));

      expect(lines).toHaveLength(2);
      expect(lines[0]?.width).toBe(width("alpha"));
      expect(lines[1]).toMatchObject({
        fromRun: 0,
        fromChar: "alpha".length + spaces.length,
      });
      expect(lines[1]?.width).toBe(width("beta"));
    }, fakeMeasure);
  });

  test("preserves leading and fitting internal spaces", () => {
    withFakeTextMeasure(() => {
      const text = "  alpha   beta";
      const { lines } = measureParagraph(paragraph([{ kind: "text", text }]), width(text));

      expect(lines).toHaveLength(1);
      expect(lines[0]).toMatchObject({ fromRun: 0, fromChar: 0, width: width(text) });
    }, fakeMeasure);
  });

  test("preserves leading spaces after an explicit line break", () => {
    withFakeTextMeasure(() => {
      const runs: Run[] = [
        { kind: "text", text: "alpha" },
        { kind: "lineBreak" },
        { kind: "text", text: "  beta" },
      ];
      const { lines } = measureParagraph(paragraph(runs), width("  beta"));

      expect(lines).toHaveLength(2);
      expect(lines[1]).toMatchObject({
        fromRun: 2,
        fromChar: 0,
        width: width("  beta"),
      });
    }, fakeMeasure);
  });

  test("keeps a split leading hyphen glued to the preceding run", () => {
    withFakeTextMeasure(() => {
      const runs: Run[] = [
        { kind: "text", text: "alpha well" },
        { kind: "text", text: "-known" },
        { kind: "text", text: " topic" },
      ];
      const maxWidth = width("alpha well") + width("-") - 0.6;
      const { lines } = measureParagraph(paragraph(runs), maxWidth);

      expect(lineStartsAt(lines, 0, "alpha ".length)).toBe(true);
      expect(lineStartsAtRun(lines, 1)).toBe(false);
    }, fakeMeasure);
  });

  test("keeps a long format-only split glued as one cluster", () => {
    withFakeTextMeasure(() => {
      const wordRuns: Run[] = "well-known".split("").map((text, index) => ({
        kind: "text",
        text,
        bold: index % 2 === 0,
      }));
      const runs: Run[] = [{ kind: "text", text: "alpha " }, ...wordRuns];
      const maxWidth = width("alpha well-") - 0.6;
      const { lines } = measureParagraph(paragraph(runs), maxWidth);

      expect(lineStartsAtRun(lines, 1)).toBe(true);
      for (let runIndex = 2; runIndex < runs.length; runIndex++) {
        expect(lineStartsAtRun(lines, runIndex)).toBe(false);
      }
    }, fakeMeasure);
  });

  test("hard-breaks oversized glued split words by line capacity", () => {
    withFakeTextMeasure(() => {
      const runs: Run[] = "abcdefghijkl".split("").map((text, index) => ({
        kind: "text",
        text,
        bold: index % 2 === 0,
      }));
      const { lines } = measureParagraph(paragraph(runs), width("abcd"));

      expect(lines[0]).toMatchObject({
        fromRun: 0,
        fromChar: 0,
        toRun: 3,
        toChar: 1,
      });
      expect(lines[1]).toMatchObject({
        fromRun: 4,
        fromChar: 0,
        toRun: 7,
        toChar: 1,
      });
      expect(lines).toHaveLength(3);
    }, fakeMeasure);
  });

  test("keeps glue when the next line can fit a cluster narrowed off the first line", () => {
    withFakeTextMeasure(() => {
      const runs: Run[] = [
        { kind: "text", text: "a " },
        { kind: "text", text: "b", bold: true },
        { kind: "text", text: "cde" },
      ];
      const { lines } = measureParagraph(
        {
          ...paragraph(runs),
          attrs: {
            listMarker: "1.",
            listMarkerSuffix: "nothing",
          },
        },
        width("abcde"),
      );

      expect(lineStartsAtRun(lines, 1)).toBe(true);
      expect(lineStartsAtRun(lines, 2)).toBe(false);
    }, fakeMeasure);
  });

  test("allows justified lines to keep a slightly overfull word", () => {
    withFakeTextMeasure(() => {
      const text = "alpha beta gamma delta";
      const runs: Run[] = [{ kind: "text", text }];
      const tightWidth = width(text) - 0.8;

      const leftMeasure = measureParagraph(paragraph(runs), tightWidth);
      const justifiedMeasure = measureParagraph(
        {
          ...paragraph(runs),
          attrs: { alignment: "justify" },
        },
        tightWidth,
      );

      expect(leftMeasure.lines).toHaveLength(2);
      expect(justifiedMeasure.lines).toHaveLength(1);
    }, fakeMeasure);
  });
});

describe("automatic hyphenation", () => {
  const paragraph = (attrs?: ParagraphBlock["attrs"]): ParagraphBlock => ({
    kind: "paragraph",
    id: "automatic-hyphenation",
    runs: [{ kind: "text", text: "hyphenation", language: { val: "en-US" } }],
    attrs,
  });

  test("uses a dictionary break and accounts for the painted hyphen width", () => {
    withFakeTextMeasure(
      () => {
        const measured = measureParagraph(
          paragraph({ automaticHyphenation: { enabled: true } }),
          70,
        );

        expect(measured.lines).toHaveLength(2);
        expect(measured.lines[0]).toMatchObject({
          fromChar: 0,
          toChar: 6,
          width: 70,
          discretionaryHyphen: { runIndex: 0 },
        });
        expect(measured.lines[1]).toMatchObject({ fromChar: 6, toChar: 11, width: 50 });
      },
      { charWidth: fixedCharWidth(10) },
    );
  });

  test("hyphenates one lexical word across formatting runs", () => {
    withFakeTextMeasure(
      () => {
        const measured = measureParagraph(
          {
            kind: "paragraph",
            id: "cross-run-automatic-hyphenation",
            runs: [
              { kind: "text", text: "xx inte", language: { val: "en-US" } },
              {
                kind: "text",
                text: "rnationalization",
                color: "#c00000",
                language: { val: "en-US" },
              },
            ],
            attrs: { automaticHyphenation: { enabled: true } },
          },
          90,
        );

        expect(measured.lines[0]).toMatchObject({
          fromRun: 0,
          fromChar: 0,
          toRun: 1,
          toChar: 1,
          width: 90,
          discretionaryHyphen: { runIndex: 1 },
        });
        expect(measured.lines[1]).toMatchObject({ fromRun: 1, fromChar: 1 });
      },
      { charWidth: fixedCharWidth(10) },
    );
  });

  test("attempts hyphenation only when line-end whitespace exceeds the document zone", () => {
    const measureWithZone = (hyphenationZoneTwips?: number) =>
      measureParagraph(
        {
          kind: "paragraph",
          id: "automatic-hyphenation-zone",
          runs: [
            {
              kind: "text",
              text: "aaaaaaaaaaaaaaaa hyphenation",
              language: { val: "en-US" },
            },
          ],
          attrs: {
            automaticHyphenation: {
              enabled: true,
              ...(hyphenationZoneTwips !== undefined ? { hyphenationZoneTwips } : {}),
            },
          },
        },
        100,
      );

    withFakeTextMeasure(
      () => {
        const defaultZone = measureWithZone();
        const zeroZone = measureWithZone(0);

        expect(defaultZone.lines[0]).toMatchObject({
          toChar: "aaaaaaaaaaaaaaaa ".length,
          width: 80,
        });
        expect(defaultZone.lines[0]?.discretionaryHyphen).toBeUndefined();
        expect(zeroZone.lines[0]).toMatchObject({
          toChar: "aaaaaaaaaaaaaaaa hy".length,
          discretionaryHyphen: { runIndex: 0 },
        });
      },
      { charWidth: fixedCharWidth(5) },
    );
  });

  test("leaves the source offsets unchanged when automatic hyphenation is suppressed", () => {
    withFakeTextMeasure(
      () => {
        const measured = measureParagraph(
          paragraph({
            automaticHyphenation: { enabled: true },
            suppressAutoHyphens: true,
          }),
          70,
        );

        expect(measured.lines[0]).toMatchObject({ fromChar: 0, toChar: 7, width: 70 });
        expect(measured.lines[0]?.discretionaryHyphen).toBeUndefined();
        expect(measured.lines[1]).toMatchObject({ fromChar: 7, toChar: 11, width: 40 });
      },
      { charWidth: fixedCharWidth(10) },
    );
  });

  test("honors the consecutive-line limit", () => {
    withFakeTextMeasure(
      () => {
        const measured = measureParagraph(
          {
            kind: "paragraph",
            id: "consecutive-hyphen-limit",
            runs: [{ kind: "text", text: "characteristically", language: { val: "en-US" } }],
            attrs: {
              automaticHyphenation: { enabled: true, consecutiveLineLimit: 1 },
            },
          },
          50,
        );

        expect(measured.lines[0]?.discretionaryHyphen).toEqual({ runIndex: 0 });
        expect(measured.lines[1]?.discretionaryHyphen).toBeUndefined();
      },
      { charWidth: fixedCharWidth(10) },
    );
  });

  test("treats the DOCX all-caps transform as capitalized text", () => {
    withFakeTextMeasure(
      () => {
        const measured = measureParagraph(
          {
            ...paragraph({
              automaticHyphenation: { enabled: true, doNotHyphenateCaps: true },
            }),
            runs: [
              {
                kind: "text",
                text: "hyphenation",
                language: { val: "en-US" },
                allCaps: true,
              },
            ],
          },
          70,
        );

        expect(measured.lines[0]?.discretionaryHyphen).toBeUndefined();
      },
      { charWidth: fixedCharWidth(10) },
    );
  });
});

describe("measureParagraph justified shrink tolerance", () => {
  const fractionalWidth = (char: string): number => (char === "b" ? 0.6 : 1);
  const ordinarySpaceRichWidth = (char: string): number => (char === "b" ? 0.4 : 1);
  const text = `${"a".repeat(99)} bbb`;
  const spaceRichText = `${"a ".repeat(20)}${"a".repeat(60)}bbb`;

  test("bases ordinary justified shrink on measured compressible spaces", () => {
    withFakeTextMeasure(
      () => {
        const spaceRichMeasure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-prose-shrink",
            runs: [{ kind: "text", text: spaceRichText }],
            attrs: { alignment: "justify" },
          },
          100,
        );
        const spacePoorMeasure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-prose-small-space-budget",
            runs: [{ kind: "text", text }],
            attrs: { alignment: "justify" },
          },
          100,
        );

        expect(spaceRichMeasure.lines).toHaveLength(1);
        expect(spacePoorMeasure.lines).toHaveLength(2);
      },
      {
        charWidth: ordinarySpaceRichWidth,
      },
    );
  });

  test("does not count non-breaking spaces toward justified shrink capacity", () => {
    const fixedSpaceText = `${"a".repeat(99)}\u00a0bbb`;

    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-fixed-space",
            runs: [{ kind: "text", text: fixedSpaceText }],
            attrs: { alignment: "justify" },
          },
          100,
        );

        expect(measure.lines).toHaveLength(2);
      },
      {
        charWidth: fractionalWidth,
      },
    );
  });

  test("keeps regular spaces available for justified shrink beside a fixed space", () => {
    const mixedSpaceText = `${"a ".repeat(15)}${"a".repeat(68)}\u00a0bbb`;

    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-mixed-spaces",
            runs: [{ kind: "text", text: mixedSpaceText }],
            attrs: { alignment: "justify" },
          },
          100,
        );

        expect(measure.lines).toHaveLength(1);
      },
      {
        charWidth: (char) => (char === "b" ? 0.5 : 1),
      },
    );
  });

  test("wraps ordinary prose beyond the constrained shrink capacity", () => {
    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-prose-shrink-boundary",
            runs: [{ kind: "text", text }],
            attrs: { alignment: "justify" },
          },
          100,
        );

        expect(measure.lines).toHaveLength(2);
      },
      {
        charWidth: (char) => (char === "b" ? 0.7 : 1),
      },
    );
  });

  test("bounds ordinary prose contraction at the measured space budget", () => {
    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-prose-space-contraction-boundary",
            runs: [{ kind: "text", text: "a a a a a a a b" }],
            attrs: { alignment: "justify" },
          },
          80,
        );

        expect(measure.lines).toHaveLength(2);
      },
      {
        charWidth: (char) => {
          if (char === " ") return 1;
          if (char === "b") return 3.6;
          return 10;
        },
      },
    );
  });

  test("ignores unused tab stops when choosing justified prose shrink tolerance", () => {
    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-unused-tab-stop",
            runs: [{ kind: "text", text: spaceRichText }],
            attrs: {
              alignment: "justify",
              indent: { firstLine: 48 },
              tabs: [{ val: "start", pos: 360 }],
            },
          },
          148,
        );

        expect(measure.lines).toHaveLength(1);
      },
      {
        charWidth: ordinarySpaceRichWidth,
      },
    );
  });

  test("keeps literal-tab first-line legal prose on the conservative shrink tolerance", () => {
    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-literal-tabbed-first-line",
            runs: [{ kind: "text", text: "x" }, { kind: "tab" }, { kind: "text", text }],
            attrs: {
              alignment: "justify",
              indent: { firstLine: 48 },
            },
          },
          100,
        );

        expect(measure.lines).toHaveLength(2);
      },
      {
        charWidth: fractionalWidth,
      },
    );
  });

  test("uses measured spaces with configured but unused tab stops", () => {
    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-hanging-tab",
            runs: [{ kind: "text", text: spaceRichText }],
            attrs: {
              alignment: "justify",
              indent: { firstLine: 0 },
              tabs: [{ val: "start", pos: 360 }],
            },
          },
          100,
        );

        expect(measure.lines).toHaveLength(1);
      },
      {
        charWidth: ordinarySpaceRichWidth,
      },
    );
  });

  test("uses the literal-tab continuation tolerance after the first line", () => {
    const continuationText = `${"a".repeat(101)}c`;
    const continuationWidth = (char: string): number => (char === "c" ? 0.65 : 1);

    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-tabbed-continuation",
            runs: [
              { kind: "text", text: "x" },
              { kind: "tab" },
              { kind: "lineBreak" },
              { kind: "text", text: continuationText },
            ],
            attrs: {
              alignment: "justify",
              indent: { firstLine: 48 },
            },
          },
          100,
        );

        expect(measure.lines).toHaveLength(2);
      },
      {
        charWidth: continuationWidth,
      },
    );
  });

  test("keeps justified list items on the conservative shrink tolerance", () => {
    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-list-item",
            runs: [{ kind: "text", text }],
            attrs: {
              alignment: "justify",
              listMarker: "•",
            },
          },
          100,
        );

        expect(measure.lines).toHaveLength(2);
      },
      {
        charWidth: fractionalWidth,
      },
    );
  });

  test("bases default list marker shrink on measured compressible spaces", () => {
    withFakeTextMeasure(
      () => {
        const spaceRichMeasure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-list-space-budget",
            runs: [{ kind: "text", text: `${"a ".repeat(15)}bbb` }],
            attrs: {
              alignment: "justify",
              indent: { left: 24, hanging: 24 },
              listMarker: "1.",
            },
          },
          53,
        );
        const spacePoorMeasure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-list-small-space-budget",
            runs: [{ kind: "text", text: `${"a".repeat(29)} bbb` }],
            attrs: {
              alignment: "justify",
              indent: { left: 24, hanging: 24 },
              listMarker: "1.",
            },
          },
          53,
        );

        expect(spaceRichMeasure.lines).toHaveLength(1);
        expect(spacePoorMeasure.lines).toHaveLength(2);
      },
      {
        charWidth: fractionalWidth,
      },
    );
  });

  test("allows the bounded first-line tolerance for deep hanging list markers", () => {
    const firstLineText = `${"a".repeat(98)} bbb`;

    withFakeTextMeasure(
      () => {
        const fittingMeasure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-deep-hanging-list-marker",
            runs: [{ kind: "text", text: firstLineText }],
            attrs: {
              alignment: "justify",
              indent: { left: 36, hanging: 36 },
              listMarker: "1.1.1",
            },
          },
          136,
        );
        const overflowingMeasure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-deep-hanging-list-marker-boundary",
            runs: [{ kind: "text", text: firstLineText }],
            attrs: {
              alignment: "justify",
              indent: { left: 36, hanging: 36 },
              listMarker: "1.1.1",
            },
          },
          135.9,
        );

        expect(fittingMeasure.lines).toHaveLength(1);
        expect(overflowingMeasure.lines).toHaveLength(2);
      },
      {
        charWidth: (char) => (char === "b" ? 1.05 : 1),
      },
    );
  });

  test("bases full-hanging list continuation shrink on measured spaces", () => {
    withFakeTextMeasure(
      () => {
        const spaceRichMeasure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-list-continuation",
            runs: [
              { kind: "text", text: "first line" },
              { kind: "lineBreak" },
              { kind: "text", text: spaceRichText },
            ],
            attrs: {
              alignment: "justify",
              listMarker: "1.",
              indent: { left: 36, hanging: 36 },
            },
          },
          136.4,
        );
        const spacePoorMeasure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-list-small-continuation-budget",
            runs: [
              { kind: "text", text: "first line" },
              { kind: "lineBreak" },
              { kind: "text", text },
            ],
            attrs: {
              alignment: "justify",
              listMarker: "1.",
              indent: { left: 36, hanging: 36 },
            },
          },
          136.4,
        );

        expect(spaceRichMeasure.lines).toHaveLength(2);
        expect(spacePoorMeasure.lines).toHaveLength(3);
      },
      {
        charWidth: fractionalWidth,
      },
    );
  });

  test("allows bounded space contraction on full-hanging list continuations", () => {
    const continuationText = `${"a ".repeat(10)}bbb`;

    withFakeTextMeasure(
      () => {
        const fittingMeasure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-list-continuation-contraction",
            runs: [
              { kind: "text", text: "first line" },
              { kind: "lineBreak" },
              { kind: "text", text: continuationText },
            ],
            attrs: {
              alignment: "justify",
              listMarker: "1.",
              indent: { left: 36, hanging: 36 },
            },
          },
          136,
        );
        const overflowingMeasure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-list-continuation-contraction-boundary",
            runs: [
              { kind: "text", text: "first line" },
              { kind: "lineBreak" },
              { kind: "text", text: continuationText },
            ],
            attrs: {
              alignment: "justify",
              listMarker: "1.",
              indent: { left: 36, hanging: 36 },
            },
          },
          135.8,
        );

        expect(fittingMeasure.lines).toHaveLength(2);
        expect(overflowingMeasure.lines).toHaveLength(3);
      },
      {
        charWidth: (char) => {
          if (char === "b") return 5.3;
          if (char === " ") return 0.55;
          return 8;
        },
      },
    );
  });

  test("uses hanging-tab shrink tolerance on a custom-hanging list marker line", () => {
    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-custom-list-marker-line",
            runs: [{ kind: "text", text }],
            attrs: {
              alignment: "justify",
              listMarker: "1.",
              indent: { left: 36, hanging: 36 },
            },
          },
          136,
        );

        expect(measure.lines).toHaveLength(1);
      },
      {
        charWidth: fractionalWidth,
      },
    );
  });

  test("does not treat non-breaking spaces as compressible list-continuation spaces", () => {
    const continuationText = `${"a".repeat(50)} ${"a".repeat(47)} \u00a0bbb`;

    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-list-nonbreaking-space",
            runs: [
              { kind: "text", text: "first line" },
              { kind: "lineBreak" },
              { kind: "text", text: continuationText },
            ],
            attrs: {
              alignment: "justify",
              listMarker: "1.",
              indent: { left: 36, hanging: 36 },
            },
          },
          136,
        );

        expect(measure.lines).toHaveLength(3);
      },
      {
        charWidth: fractionalWidth,
      },
    );
  });

  test("uses the measured-space budget for full-hanging list continuations", () => {
    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-standard-list-continuation",
            runs: [
              { kind: "text", text: "first line" },
              { kind: "lineBreak" },
              { kind: "text", text: spaceRichText },
            ],
            attrs: {
              alignment: "justify",
              listMarker: "1.",
              indent: { left: 24, hanging: 24 },
            },
          },
          124.4,
        );

        expect(measure.lines).toHaveLength(2);
      },
      {
        charWidth: fractionalWidth,
      },
    );
  });

  test("counts spaces inside opaque field runs but not native math", () => {
    const opaqueSpaceRichText = `${"a ".repeat(10)}b`;
    const opaqueSpacePoorText = `${"a".repeat(19)} b`;
    const measureOpaqueRun = (run: Run) =>
      measureParagraph(
        {
          kind: "paragraph",
          id: "justified-opaque-run-space-budget",
          runs: [{ kind: "text", text: "x" }, run],
          attrs: { alignment: "justify" },
        },
        21.3,
      );
    const measureResolvedField = () =>
      measureParagraph(
        {
          kind: "paragraph",
          id: "justified-resolved-field-space-budget",
          runs: [
            { kind: "text", text: "x" },
            {
              kind: "field",
              fieldType: "OTHER",
              fallback: opaqueSpacePoorText,
              pmStart: 1,
            },
          ],
          attrs: { alignment: "justify" },
        },
        21.3,
        { fieldValues: new Map([[1, opaqueSpaceRichText]]) },
      );

    withFakeTextMeasure(
      () => {
        const fieldRich = measureOpaqueRun({
          kind: "field",
          fieldType: "OTHER",
          fallback: opaqueSpaceRichText,
        });
        const fieldPoor = measureOpaqueRun({
          kind: "field",
          fieldType: "OTHER",
          fallback: opaqueSpacePoorText,
        });
        const mathRich = measureOpaqueRun({
          kind: "math",
          display: "inline",
          ommlXml:
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>a b</m:t></m:r></m:oMath>',
          plainText: opaqueSpaceRichText,
        });

        expect(fieldRich.lines).toHaveLength(1);
        expect(fieldPoor.lines).toHaveLength(2);
        expect(measureResolvedField().lines).toHaveLength(1);
        expect(mathRich.lines).toHaveLength(2);
      },
      { charWidth: fixedCharWidth(1) },
    );
  });

  test("keeps shallow full-hanging list continuations on rounding tolerance", () => {
    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-shallow-list-continuation",
            runs: [
              { kind: "text", text: "first line" },
              { kind: "lineBreak" },
              { kind: "text", text },
            ],
            attrs: {
              alignment: "justify",
              listMarker: "1.",
              indent: { left: 18, hanging: 18 },
            },
          },
          118,
        );

        expect(measure.lines).toHaveLength(3);
      },
      {
        charWidth: fractionalWidth,
      },
    );
  });

  test("keeps inset list continuation lines on the conservative tolerance", () => {
    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-inset-list-continuation",
            runs: [
              { kind: "text", text: "first line" },
              { kind: "lineBreak" },
              { kind: "text", text },
            ],
            attrs: {
              alignment: "justify",
              listMarker: "1.",
              indent: { left: 36, hanging: 18 },
            },
          },
          136,
        );

        expect(measure.lines).toHaveLength(3);
      },
      {
        charWidth: fractionalWidth,
      },
    );
  });

  test("wraps a trailing token beyond the inset-list continuation allowance", () => {
    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "justified-inset-list-boundary",
            runs: [
              { kind: "text", text: "x" },
              { kind: "lineBreak" },
              { kind: "text", text: `${"a".repeat(89)} bbb` },
            ],
            attrs: {
              alignment: "justify",
              listMarker: "1.",
              indent: { left: 20, hanging: 10 },
            },
          },
          110,
        );

        expect(measure.lines).toHaveLength(3);
      },
      {
        charWidth: (char) => (char === "b" ? 0.465 : 1),
      },
    );
  });
});

describe("measureParagraph document default tab interval", () => {
  test("advances consecutive tabs on the authored grid", () => {
    withFakeTextMeasure(
      () => {
        const measure = measureParagraph(
          {
            kind: "paragraph",
            id: "authored-default-tabs",
            runs: [
              { kind: "text", text: "a".repeat(41) },
              { kind: "tab" },
              { kind: "tab" },
              { kind: "text", text: "x" },
            ],
            attrs: { defaultTabStopTwips: 600 },
          },
          200,
        );

        expect(measure.lines).toHaveLength(1);
        expect(measure.lines[0]?.width).toBe(121);
      },
      { charWidth: () => 1 },
    );
  });
});

describe("inline image paragraph measurement", () => {
  test("image-only line uses the authored image footprint", () => {
    const imageHeight = 29;
    const measure = measureParagraph(
      {
        kind: "paragraph",
        id: "img1",
        pmStart: 0,
        pmEnd: 1,
        runs: [
          {
            kind: "image",
            src: "data:image/png;base64,",
            width: 186,
            height: imageHeight,
            pmStart: 0,
            pmEnd: 1,
          },
        ],
        attrs: {
          defaultFontSize: 11,
          defaultFontFamily: "Calibri",
        },
      },
      600,
    );

    expect(measure.lines[0]?.lineHeight).toBe(imageHeight);
    expect(measure.lines[0]?.ascent).toBe(imageHeight);
    expect(measure.lines[0]?.descent).toBe(0);
  });

  test("advances floating-zone y offsets by image-inflated line height", () => {
    withFakeTextMeasure(() => {
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "img-float",
          runs: [
            {
              kind: "image",
              src: "data:image/png;base64,",
              width: 35,
              height: 90,
            },
            { kind: "text", text: "iiii" },
          ],
        },
        40,
        {
          floatingZones: [
            {
              leftMargin: 35,
              rightMargin: 0,
              topY: 20,
              bottomY: 80,
            },
          ],
        },
      );

      expect(measure.lines).toHaveLength(2);
      expect(measure.lines.at(1)?.leftOffset).toBeUndefined();
      expect(measure.lines.at(1)?.width).toBe(20);
    }, fakeMeasure);
  });

  // Regression: a logo + label header line (image flowing alongside text) used
  // to inherit the image-only branch's `imageH + descent*2` line box, which
  // centered the text inside an inflated band and left it floating above the
  // paragraph border (eigenpal #580). Word baseline-aligns the row and sizes
  // the line as `imageH + text descent`.
  test("image-with-text line sizes as imageH + text descent (baseline-aligned)", () => {
    withFakeTextMeasure(() => {
      const imageHeight = 40;
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "logo-label",
          runs: [
            {
              kind: "image",
              src: "data:image/png;base64,",
              width: 80,
              height: imageHeight,
            },
            { kind: "text", text: "Header" },
          ],
        },
        600,
      );

      const line = measure.lines.at(0);
      expect(line).toBeDefined();
      // Line height should equal imageH + a single descent buffer; the
      // image-alone branch would emit imageH + descent*2, so check that the
      // height is strictly less than that.
      const descent = line?.descent ?? 0;
      expect(descent).toBeGreaterThan(0);
      expect(line?.lineHeight).toBe(imageHeight + descent);
      expect(line?.ascent).toBe(imageHeight);
    }, fakeMeasure);
  });

  test("image-only line does not add text descent", () => {
    const imageHeight = 40;
    const measure = measureParagraph(
      {
        kind: "paragraph",
        id: "logo-alone",
        runs: [
          {
            kind: "image",
            src: "data:image/png;base64,",
            width: 80,
            height: imageHeight,
          },
        ],
        attrs: {
          defaultFontSize: 11,
          defaultFontFamily: "Calibri",
        },
      },
      600,
    );

    const line = measure.lines.at(0);
    expect(line).toBeDefined();
    expect(line?.lineHeight).toBe(imageHeight);
    expect(line?.ascent).toBe(imageHeight);
    expect(line?.descent).toBe(0);
  });

  test("embedded-object preview uses its authored box as the exact line height", () => {
    const imageHeight = 40;
    const measure = measureParagraph(
      {
        kind: "paragraph",
        id: "object-preview",
        runs: [
          {
            kind: "image",
            src: "data:image/png;base64,",
            width: 80,
            height: imageHeight,
            exactLineHeight: true,
          },
        ],
        attrs: {
          defaultFontSize: 11,
          defaultFontFamily: "Calibri",
        },
      },
      600,
    );

    expect(measure.lines[0]).toMatchObject({
      lineHeight: imageHeight,
      ascent: imageHeight,
      descent: 0,
    });
    expect(measure.totalHeight).toBe(imageHeight);
  });

  test("inline image footprint includes its wp:inline distT/distB", () => {
    withFakeTextMeasure(() => {
      // distTop/distBottom = 8 each = 16px of extra footprint; the line
      // height must reserve that or the painter's per-image margin spills
      // past the line's reserved height.
      const imageHeight = 20;
      const distTop = 8;
      const distBottom = 8;
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "logo-dist",
          runs: [
            {
              kind: "image",
              src: "data:image/png;base64,",
              width: 80,
              height: imageHeight,
              distTop,
              distBottom,
            },
            { kind: "text", text: "Header" },
          ],
        },
        600,
      );

      const line = measure.lines.at(0);
      expect(line).toBeDefined();
      const descent = line?.descent ?? 0;
      // Footprint = imageHeight + distTop + distBottom; the image-with-text
      // branch adds a single descent buffer below baseline.
      expect(line?.lineHeight).toBe(imageHeight + distTop + distBottom + descent);
    }, fakeMeasure);
  });

  // Regression (eigenpal/docx-editor#767, fixes #766): an inline image that
  // wraps to its own line must inflate only the line it lands on. The footprint
  // was recorded on the current line *before* the wrap check, so the preceding
  // text line inflated to image height while the image's own line stayed
  // text-height — and following content painted over the overflowing image.
  test("inline image that wraps reserves its height on its own line, not the text line", () => {
    withFakeTextMeasure(() => {
      const imageHeight = 151;
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "wrap-img",
          runs: [
            { kind: "text", text: "Some preceding text" },
            // One px shy of the column, so any preceding text forces the image
            // onto its own line where it still fits.
            {
              kind: "image",
              src: "data:image/png;base64,",
              width: 599,
              height: imageHeight,
            },
            { kind: "text", text: "Description" },
          ],
        },
        600,
      );

      expect(measure.lines.length).toBeGreaterThanOrEqual(2);
      // The preceding text line stays text-height, not inflated to the image.
      expect(measure.lines.at(0)?.lineHeight ?? 0).toBeLessThan(imageHeight);
      // The image lands on the next line, which reserves its full height.
      expect(measure.lines.at(1)?.lineHeight ?? 0).toBeGreaterThanOrEqual(imageHeight);
    }, fakeMeasure);
  });

  test("inline image that fits inline still grows its shared line to image height", () => {
    withFakeTextMeasure(() => {
      const imageHeight = 151;
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "fit-img",
          runs: [
            { kind: "text", text: "Hi" },
            {
              kind: "image",
              src: "data:image/png;base64,",
              width: 100,
              height: imageHeight,
            },
          ],
        },
        600,
      );

      expect(measure.lines).toHaveLength(1);
      expect(measure.lines.at(0)?.lineHeight ?? 0).toBeGreaterThanOrEqual(imageHeight);
    }, fakeMeasure);
  });

  // An inline image that is the only run on its line stays on that line even
  // when it is wider than the column, instead of being wrapped to a fresh blank
  // row above it — wrapping an image that is already alone can't make it fit.
  // The measurer reserves its intrinsic box (the painter caps the paint with
  // CSS max-width:100%; see measureParagraph.ts). (eigenpal/docx-editor#760.)
  test("inline image wider than the column stays on its line and reserves its full height", () => {
    withFakeTextMeasure(() => {
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "wide-img",
          runs: [
            {
              kind: "image",
              src: "data:image/png;base64,",
              width: 1200,
              height: 800,
            },
          ],
        },
        600,
      );

      // No spurious leading empty line; the image is the only line.
      expect(measure.lines).toHaveLength(1);
      // The intrinsic height (800) is reserved, never under-reserved.
      expect(measure.lines.at(0)?.lineHeight ?? 0).toBeGreaterThanOrEqual(800);
    }, fakeMeasure);
  });

  // Regression: a 100×200 inline image rotated 90° should reserve a 200×100
  // axis-aligned bbox in the measurer, matching the painter's wrapper span
  // (eigenpal #424 follow-up; gemini/codex review on PR 518). Without this,
  // following text wrapped too early horizontally and the next line could
  // overlap vertically.
  test("rotated inline image reserves its bbox width on the line", () => {
    const measure = measureParagraph(
      {
        kind: "paragraph",
        id: "img-rot-w",
        runs: [
          {
            kind: "image",
            src: "data:image/png;base64,",
            width: 100,
            height: 200,
            transform: "rotate(90deg)",
          },
        ],
      },
      600,
    );

    const line = measure.lines.at(0);
    expect(line).toBeDefined();
    // bbox width = 200 (swapped from 100×200). The raw run.width would be 100.
    expect(line?.width).toBe(200);
  });

  test("rotated inline image reserves its bbox height on the line", () => {
    const imageWidth = 200;
    const imageHeight = 100;
    const measure = measureParagraph(
      {
        kind: "paragraph",
        id: "img-rot-h",
        runs: [
          {
            kind: "image",
            src: "data:image/png;base64,",
            width: imageWidth,
            height: imageHeight,
            transform: "rotate(90deg)",
          },
        ],
        attrs: {
          defaultFontSize: 11,
          defaultFontFamily: "Calibri",
        },
      },
      600,
    );

    const line = measure.lines.at(0);
    expect(line).toBeDefined();
    // bbox height = imageWidth (swapped). With the previous bug the line
    // would reserve `imageHeight` (100), which is shorter than the painted
    // bbox (200) and would let the next line overlap the rotated picture.
    expect(line?.ascent).toBeGreaterThanOrEqual(imageWidth);
  });

  test("rotated portrait→landscape inline image wraps onto a new line when bbox exceeds availableWidth", () => {
    withFakeTextMeasure(() => {
      // Container width 150; raw run.width = 100 would fit, but the rotated
      // bbox width = 200 should force a wrap onto its own line.
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "img-rot-wrap",
          runs: [
            { kind: "text", text: "x" },
            {
              kind: "image",
              src: "data:image/png;base64,",
              width: 100,
              height: 200,
              transform: "rotate(90deg)",
            },
          ],
        },
        150,
      );

      // Two lines: leading text on one, the rotated image alone on the next.
      expect(measure.lines.length).toBeGreaterThanOrEqual(2);
      expect(measure.lines.at(-1)?.width).toBe(200);
    }, fakeMeasure);
  });

  test("inline image with no rotation keeps raw width on the line", () => {
    const measure = measureParagraph(
      {
        kind: "paragraph",
        id: "img-no-rot",
        runs: [
          {
            kind: "image",
            src: "data:image/png;base64,",
            width: 100,
            height: 200,
          },
        ],
      },
      600,
    );

    const line = measure.lines.at(0);
    expect(line).toBeDefined();
    expect(line?.width).toBe(100);
  });
});

describe("math paragraph measurement", () => {
  test("stacked inline math reserves extra line height", () => {
    withFakeTextMeasure(() => {
      const inlineMeasure = measureParagraph(
        {
          kind: "paragraph",
          id: "math-inline",
          runs: [
            {
              kind: "math",
              display: "inline",
              ommlXml: "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>",
              plainText: "x",
              fontSize: 11,
            },
          ],
        },
        600,
      );
      const stackedMeasure = measureParagraph(
        {
          kind: "paragraph",
          id: "math-stacked",
          runs: [
            {
              kind: "math",
              display: "inline",
              ommlXml:
                "<m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>",
              plainText: "a/b",
              fontSize: 11,
            },
          ],
        },
        600,
      );

      const inlineLineHeight = inlineMeasure.lines.at(0)?.lineHeight ?? 0;
      const stackedLineHeight = stackedMeasure.lines.at(0)?.lineHeight ?? 0;
      expect(stackedLineHeight).toBeGreaterThan(inlineLineHeight + 10);
    }, fakeMeasure);
  });
});

describe("block image rotation measurement", () => {
  // The painter wraps a rotated block image in an axis-aligned bbox
  // (`renderBlockImage`, eigenpal #424). The measurer has to reserve the
  // same rotated bbox height; otherwise the painter's container overflows
  // the line box and the next paragraph paints on top of the rotated
  // landscape image (codex PR #521 review).
  test("rotated block image reserves the rotated bbox height (270deg landscape)", () => {
    const measure = measureParagraph(
      {
        kind: "paragraph",
        id: "rot-block",
        pmStart: 0,
        pmEnd: 1,
        runs: [
          {
            kind: "image",
            src: "data:image/png;base64,",
            width: 120,
            height: 60,
            displayMode: "block",
            transform: "rotate(270deg)",
            distTop: 6,
            distBottom: 6,
          },
        ],
        attrs: {
          defaultFontSize: 11,
          defaultFontFamily: "Calibri",
        },
      },
      600,
    );

    // Rotated bbox of 120x60 at 270deg is 60x120: reserve the 120px
    // rotated height plus the default 6+6 margins, not the intrinsic
    // 60px height the un-rotated path used.
    expect(measure.lines[0]?.lineHeight).toBeGreaterThanOrEqual(120 + 12);
  });

  test("un-rotated block image still reserves the intrinsic height", () => {
    const measure = measureParagraph(
      {
        kind: "paragraph",
        id: "noop-block",
        pmStart: 0,
        pmEnd: 1,
        runs: [
          {
            kind: "image",
            src: "data:image/png;base64,",
            width: 120,
            height: 60,
            displayMode: "block",
            distTop: 6,
            distBottom: 6,
          },
        ],
        attrs: {
          defaultFontSize: 11,
          defaultFontFamily: "Calibri",
        },
      },
      600,
    );

    expect(measure.lines[0]?.lineHeight).toBeGreaterThanOrEqual(60 + 12);
    expect(measure.lines[0]?.lineHeight).toBeLessThan(120);
  });

  test("positioned top-and-bottom artwork stays out of paragraph flow", () => {
    withFakeTextMeasure(() => {
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "positioned-band",
          runs: [
            {
              kind: "image",
              src: "data:image/png;base64,",
              width: 600,
              height: 95,
              displayMode: "block",
              wrapType: "topAndBottom",
              position: {
                horizontal: { relativeTo: "page", posOffset: 0 },
                vertical: { relativeTo: "page", posOffset: 0 },
              },
            },
            { kind: "text", text: "Body text" },
          ],
        },
        600,
      );

      expect(measure.lines).toHaveLength(1);
      expect(measure.lines.at(0)?.lineHeight).toBeLessThan(95);
    }, fakeMeasure);
  });
});

describe("paragraph indentation measurement", () => {
  test("negative side indents widen the measured line box", () => {
    withFakeTextMeasure(() => {
      const block = {
        kind: "paragraph" as const,
        id: "negative-indent",
        runs: [
          {
            kind: "text" as const,
            text: "aaaaaaaaaaaaaaaaaaaaaaaaaa",
            fontSize: 11,
          },
        ],
        attrs: {
          indent: {
            left: -20,
            right: -10,
          },
        },
      };

      expect(measureParagraph(block, 100).lines).toHaveLength(1);
    }, fakeMeasure);
  });
});

describe("all-caps paragraph measurement", () => {
  test("builds canvas font strings with the rendered DOCX bold weight", () => {
    const font = buildFontString({
      fontFamily: "Arial",
      fontSize: 12,
      bold: true,
    });

    expect(font).toContain("800");
    expect(font).not.toContain("bold");
  });

  test("measures all-caps runs using uppercase glyph widths", () => {
    withFakeTextMeasure(() => {
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "caps",
          runs: [{ kind: "text", text: "iiii", allCaps: true }],
        },
        25,
      );

      expect(measure.lines).toHaveLength(2);
    }, fakeMeasure);
  });

  test("measures horizontally scaled runs using scaled glyph widths", () => {
    withFakeTextMeasure(() => {
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "scaled",
          runs: [{ kind: "text", text: "iiii", horizontalScale: 150 }],
        },
        25,
      );

      expect(measure.lines).toHaveLength(2);
    }, fakeMeasure);
  });

  test("measures small-caps runs using small-caps glyph widths", () => {
    withFakeTextMeasure(() => {
      const measure = measureParagraph(
        {
          kind: "paragraph",
          id: "small-caps",
          runs: [{ kind: "text", text: "iiii", smallCaps: true }],
        },
        25,
      );

      expect(measure.lines).toHaveLength(2);
      expect(getRunCharWidths({ kind: "text", text: "ii", smallCaps: true })).toEqual([8, 8]);
    }, fakeMeasure);
  });

  test("includes small-caps formatting in paragraph cache keys", () => {
    const plainHash = hashParagraphBlock({
      kind: "paragraph",
      id: "plain",
      runs: [{ kind: "text", text: "iiii" }],
    });
    const smallCapsHash = hashParagraphBlock({
      kind: "paragraph",
      id: "small-caps",
      runs: [{ kind: "text", text: "iiii", smallCaps: true }],
    });

    expect(smallCapsHash).not.toBe(plainHash);
  });

  test("includes character spacing in paragraph cache keys", () => {
    const plainHash = hashParagraphBlock({
      kind: "paragraph",
      id: "plain",
      runs: [{ kind: "text", text: "iiii" }],
    });
    const spacedHash = hashParagraphBlock({
      kind: "paragraph",
      id: "spaced",
      runs: [{ kind: "text", text: "iiii", letterSpacing: 1.5 }],
    });

    expect(spacedHash).not.toBe(plainHash);
  });

  test("includes the East Asian font in paragraph cache keys", () => {
    // Measurement now depends on eastAsiaFontFamily, so the cache key must too —
    // otherwise a CJK run keeps stale measurements when only its EA face changes.
    const minchoHash = hashParagraphBlock({
      kind: "paragraph",
      id: "mincho",
      runs: [{ kind: "text", text: "日本語", eastAsiaFontFamily: "MS Mincho" }],
    });
    const gothicHash = hashParagraphBlock({
      kind: "paragraph",
      id: "gothic",
      runs: [{ kind: "text", text: "日本語", eastAsiaFontFamily: "MS Gothic" }],
    });

    expect(minchoHash).not.toBe(gothicHash);
  });
});

describe("CJK line breaking", () => {
  test("fills the remaining width on a line before wrapping CJK characters", () => {
    withFakeTextMeasure(
      () => {
        const block: ParagraphBlock = {
          kind: "paragraph",
          id: "cjk-wrap",
          pmStart: 0,
          pmEnd: 8,
          runs: [
            { kind: "text", text: "AAAA" },
            { kind: "text", text: "一二三四", eastAsiaFontFamily: "MS Mincho" },
          ],
        };

        const measure = measureParagraph(block, 50);

        expect(measure.lines).toHaveLength(2);
        expect(measure.lines[0]?.toRun).toBe(1);
        expect(measure.lines[0]?.toChar).toBe(1);
        expect(measure.lines[1]?.fromRun).toBe(1);
        expect(measure.lines[1]?.fromChar).toBe(1);
        expect(measure.lines[1]?.toChar).toBe(4);
      },
      { charWidth: fixedCharWidth(10) },
    );
  });

  test("does not wrap preceding Latin text when the next CJK character does not fit", () => {
    withFakeTextMeasure(
      () => {
        const block: ParagraphBlock = {
          kind: "paragraph",
          id: "cjk-wrap-boundary",
          pmStart: 0,
          pmEnd: 8,
          runs: [
            { kind: "text", text: "AAAA" },
            { kind: "text", text: "一二三四", eastAsiaFontFamily: "MS Mincho" },
          ],
        };

        const measure = measureParagraph(block, 45);

        expect(measure.lines).toHaveLength(2);
        expect(measure.lines[0]?.toRun).toBe(0);
        expect(measure.lines[0]?.toChar).toBe(4);
        expect(measure.lines[1]?.fromRun).toBe(1);
        expect(measure.lines[1]?.fromChar).toBe(0);
      },
      { charWidth: fixedCharWidth(10) },
    );
  });

  test("defaults to hanging punctuation unless w:overflowPunct is disabled", () => {
    withFakeTextMeasure(
      () => {
        const paragraph = (overflowPunctuation?: boolean): ParagraphBlock => ({
          kind: "paragraph",
          id: `cjk-overflow-punctuation-${String(overflowPunctuation)}`,
          runs: [{ kind: "text", text: "中文。", language: { eastAsia: "zh-CN" } }],
          ...(overflowPunctuation === undefined ? {} : { attrs: { overflowPunctuation } }),
        });

        expect(measureParagraph(paragraph(false), 20).lines).toHaveLength(2);
        for (const hanging of [
          measureParagraph(paragraph(), 20),
          measureParagraph(paragraph(true), 20),
        ]) {
          expect(hanging.lines).toHaveLength(1);
          expect(hanging.lines[0]?.width).toBe(30);
        }
      },
      { charWidth: fixedCharWidth(10) },
    );
  });

  test("does not hang opening punctuation", () => {
    withFakeTextMeasure(
      () => {
        const paragraph: ParagraphBlock = {
          kind: "paragraph",
          id: "cjk-opening-punctuation",
          runs: [{ kind: "text", text: "中文（", language: { eastAsia: "zh-CN" } }],
          attrs: { overflowPunctuation: true },
        };

        expect(measureParagraph(paragraph, 20).lines.length).toBeGreaterThan(1);
      },
      { charWidth: fixedCharWidth(10) },
    );
  });

  test("hangs a document-specific prohibited line-start character", () => {
    withFakeTextMeasure(
      () => {
        const paragraph: ParagraphBlock = {
          kind: "paragraph",
          id: "custom-hanging-punctuation",
          runs: [{ kind: "text", text: "中文※", language: { eastAsia: "zh-CN" } }],
          attrs: {
            overflowPunctuation: true,
            lineBreakRules: {
              noLineBreaksBefore: { language: "zh-CN", characters: "※" },
            },
          },
        };

        expect(measureParagraph(paragraph, 20).lines).toHaveLength(1);
      },
      { charWidth: fixedCharWidth(10) },
    );
  });
});

describe("CJK line height", () => {
  // Word derives a CJK line's height from an East-Asian face, not the run's
  // ascii font: runs whose ascii/eastAsia fonts are Latin (common in real
  // Japanese documents, e.g. w:eastAsia="Century") still render at the default
  // East-Asian face's ≈1.303 single-line ratio, not the Latin ≈1.15. All
  // fixtures use the default 11pt size; no explicit spacing, so
  // lineHeight = fontSizePx × singleLineRatio.
  const fontSizePx = 11 * PT_TO_PX;
  const latinLineHeight = fontSizePx * 1.15; // unmapped "Century" → default ratio
  const cjkLineHeight = fontSizePx * 1.303; // measured East-Asian ratio

  const firstLineHeight = (runs: Run[]): number => {
    const measure = measureParagraph(
      { kind: "paragraph", id: "cjk-height", pmStart: 0, pmEnd: 1, runs },
      600,
    );
    return measure.lines[0]?.lineHeight ?? 0;
  };

  test("CJK text with a Latin eastAsia font uses the East-Asian fallback ratio", () => {
    withFakeTextMeasure(() => {
      const lineHeight = firstLineHeight([
        {
          kind: "text",
          text: "秘密保持契約",
          fontFamily: "Century",
          eastAsiaFontFamily: "Century",
        },
      ]);

      expect(lineHeight).toBeCloseTo(cjkLineHeight, 2);
    }, fakeMeasure);
  });

  test("CJK text with no eastAsia font uses the East-Asian fallback ratio", () => {
    withFakeTextMeasure(() => {
      const lineHeight = firstLineHeight([
        { kind: "text", text: "秘密保持契約", fontFamily: "Century" },
      ]);

      expect(lineHeight).toBeCloseTo(cjkLineHeight, 2);
    }, fakeMeasure);
  });

  test("CJK text with a real CJK eastAsia font uses that font's own ratio", () => {
    withFakeTextMeasure(() => {
      // SimSun keeps the 1.15 default ratio — landing there (not at the 1.303
      // fallback) proves the eastAsia face itself supplied the line height.
      const simsunHeight = firstLineHeight([
        { kind: "text", text: "保密协议", fontFamily: "Century", eastAsiaFontFamily: "SimSun" },
      ]);
      const minchoHeight = firstLineHeight([
        {
          kind: "text",
          text: "秘密保持契約",
          fontFamily: "Century",
          eastAsiaFontFamily: "MS Mincho",
        },
      ]);

      expect(simsunHeight).toBeCloseTo(fontSizePx * 1.15, 2);
      expect(minchoHeight).toBeCloseTo(cjkLineHeight, 2);
    }, fakeMeasure);
  });

  test("pure Latin runs keep the ascii font's ratio", () => {
    withFakeTextMeasure(() => {
      const lineHeight = firstLineHeight([
        {
          kind: "text",
          text: "Confidential Agreement",
          fontFamily: "Century",
          eastAsiaFontFamily: "Century",
        },
      ]);

      expect(lineHeight).toBeCloseTo(latinLineHeight, 2);
    }, fakeMeasure);
  });

  test("mixed Latin + CJK line takes the taller CJK height, in either run order", () => {
    withFakeTextMeasure(() => {
      const latinFirst = firstLineHeight([
        { kind: "text", text: "Article 1 ", fontFamily: "Century" },
        { kind: "text", text: "秘密保持", fontFamily: "Century", eastAsiaFontFamily: "Century" },
      ]);
      const cjkFirst = firstLineHeight([
        { kind: "text", text: "秘密保持", fontFamily: "Century", eastAsiaFontFamily: "Century" },
        { kind: "text", text: " (Confidentiality)", fontFamily: "Century" },
      ]);

      expect(latinFirst).toBeCloseTo(cjkLineHeight, 2);
      expect(cjkFirst).toBeCloseTo(cjkLineHeight, 2);
    }, fakeMeasure);
  });

  test("a larger Latin font still dominates a smaller CJK run (size precedence preserved)", () => {
    withFakeTextMeasure(() => {
      const lineHeight = firstLineHeight([
        { kind: "text", text: "Heading ", fontFamily: "Century", fontSize: 16 },
        { kind: "text", text: "見出し", fontFamily: "Century", fontSize: 11 },
      ]);

      expect(lineHeight).toBeCloseTo(16 * PT_TO_PX * 1.15, 2);
    }, fakeMeasure);
  });
});

describe("clampFloatingWrapMargins", () => {
  // A near-full-width floating table or image computes a left/right wrap
  // margin that extends past contentWidth (margins are `rectRight` or
  // `contentWidth - (x - distLeft)`, both of which can spill). Without
  // clamping, getFloatingMargins propagates that margin into the line and
  // measureParagraph collapses every wrapped line to ~1 glyph wide — the
  // "single character per line after a wide float" symptom.
  test("zeros margins that exceed content width", () => {
    expect(clampFloatingWrapMargins(698, 0, 671)).toEqual({
      leftMargin: 0,
      rightMargin: 0,
    });
    expect(clampFloatingWrapMargins(0, 700, 671)).toEqual({
      leftMargin: 0,
      rightMargin: 0,
    });
  });

  test("zeros when combined side margins cover the content area", () => {
    expect(clampFloatingWrapMargins(400, 300, 671)).toEqual({
      leftMargin: 0,
      rightMargin: 0,
    });
  });

  test("preserves valid one-sided margins", () => {
    expect(clampFloatingWrapMargins(200, 0, 671)).toEqual({
      leftMargin: 200,
      rightMargin: 0,
    });
    expect(clampFloatingWrapMargins(0, 150, 671)).toEqual({
      leftMargin: 0,
      rightMargin: 150,
    });
  });

  test("clamps negative inputs to 0", () => {
    expect(clampFloatingWrapMargins(-5, -10, 671)).toEqual({
      leftMargin: 0,
      rightMargin: 0,
    });
  });

  test("falls back to contentWidth=1 floor for non-positive contentWidth", () => {
    expect(clampFloatingWrapMargins(0, 0, 0)).toEqual({
      leftMargin: 0,
      rightMargin: 0,
    });
  });
});
