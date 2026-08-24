import { expect, test } from "bun:test";

import type { ParagraphBlock } from "../types";
import { fixedCharWidth, withFakeTextMeasure } from "./__tests__/fakeTextMeasure";
import {
  clearAllCaches,
  clearTextWidthCache,
  getCachedTextWidth,
  getTextCacheSize,
  hashParagraphBlock,
  setCachedTextWidth,
  setTextCacheSize,
} from "./cache";
import { measureBlock } from "./measureBlocks";

const measureParagraphBlock = (block: ParagraphBlock, width: number) => {
  const measure = measureBlock(block, width);
  if (measure.kind !== "paragraph") {
    throw new Error("expected paragraph measure");
  }
  return measure;
};

const expectCachedMeasurementsMatchFreshInBothOrders = (
  first: ParagraphBlock,
  second: ParagraphBlock,
  width: number,
): void => {
  clearAllCaches();
  const freshFirst = measureParagraphBlock(first, width);
  clearAllCaches();
  const freshSecond = measureParagraphBlock(second, width);
  expect(freshFirst).not.toEqual(freshSecond);

  clearAllCaches();
  expect(measureParagraphBlock(first, width)).toEqual(freshFirst);
  expect(measureParagraphBlock(second, width)).toEqual(freshSecond);

  clearAllCaches();
  expect(measureParagraphBlock(second, width)).toEqual(freshSecond);
  expect(measureParagraphBlock(first, width)).toEqual(freshFirst);
};

test("text width cache retains frequently reused entries when it evicts", () => {
  try {
    setTextCacheSize(2);
    clearTextWidthCache();
    setCachedTextWidth("hot", "16px Arial", 0, 1);
    setCachedTextWidth("cold", "16px Arial", 0, 2);

    for (let hit = 0; hit < 32; hit += 1) {
      expect(getCachedTextWidth("hot", "16px Arial", 0)).toBe(1);
    }
    setCachedTextWidth("new", "16px Arial", 0, 3);

    expect(getCachedTextWidth("hot", "16px Arial", 0)).toBe(1);
    expect(getCachedTextWidth("cold", "16px Arial", 0)).toBeUndefined();
    expect(getCachedTextWidth("new", "16px Arial", 0)).toBe(3);
    expect(getTextCacheSize()).toBe(2);
  } finally {
    setTextCacheSize(20_000);
    clearTextWidthCache();
  }
});

test("paragraph cache keys include complex-script list marker formatting", () => {
  const paragraphWithMarkerSize = (complexScriptFontSize: number): ParagraphBlock => ({
    kind: "paragraph",
    id: "list-item",
    runs: [{ kind: "text", text: "بند" }],
    attrs: {
      listMarker: "ا.",
      listMarkerFormatting: {
        complexScriptFontFamily: "Traditional Arabic",
        complexScriptFontSize,
      },
    },
  });

  expect(hashParagraphBlock(paragraphWithMarkerSize(12))).not.toBe(
    hashParagraphBlock(paragraphWithMarkerSize(18)),
  );
});

test("paragraph cache keys include list paragraph mark font size", () => {
  const paragraph = {
    kind: "paragraph",
    id: "list-paragraph-mark",
    runs: [{ kind: "text", text: "aaaa aaaa", fontFamily: "Aptos", fontSize: 10.5 }],
    attrs: {
      defaultFontSize: 10.5,
      defaultFontFamily: "Aptos",
      listMarker: "\u2022",
      listMarkerSuffix: "nothing",
      listParagraphMarkFontSize: 11,
    },
  } as const satisfies ParagraphBlock;

  expect(hashParagraphBlock(paragraph)).not.toBe(
    hashParagraphBlock({
      ...paragraph,
      attrs: { ...paragraph.attrs, listParagraphMarkFontSize: 18 },
    }),
  );
});

test("paragraph cache keys include every tab layout input", () => {
  const paragraph = {
    kind: "paragraph",
    id: "tab-layout",
    runs: [{ kind: "text", text: "A" }, { kind: "tab" }, { kind: "text", text: "B" }],
    attrs: {
      bidi: true,
      defaultTabStopTwips: 720,
      tabs: [
        { val: "start", pos: 1500, leader: "dot" },
        { val: "end", pos: 3000, leader: "hyphen" },
      ],
    },
  } as const satisfies ParagraphBlock;
  const hash = hashParagraphBlock(paragraph);

  expect(
    hashParagraphBlock({
      ...paragraph,
      attrs: { ...paragraph.attrs, bidi: false },
    }),
  ).not.toBe(hash);
  expect(
    hashParagraphBlock({
      ...paragraph,
      attrs: { ...paragraph.attrs, defaultTabStopTwips: 1440 },
    }),
  ).not.toBe(hash);
  expect(
    hashParagraphBlock({
      ...paragraph,
      attrs: {
        ...paragraph.attrs,
        tabs: [
          { val: "start", pos: 1501, leader: "dot" },
          { val: "end", pos: 3000, leader: "hyphen" },
        ],
      },
    }),
  ).not.toBe(hash);
  expect(
    hashParagraphBlock({
      ...paragraph,
      attrs: {
        ...paragraph.attrs,
        tabs: [
          { val: "center", pos: 1500, leader: "dot" },
          { val: "end", pos: 3000, leader: "hyphen" },
        ],
      },
    }),
  ).not.toBe(hash);
  expect(
    hashParagraphBlock({
      ...paragraph,
      attrs: {
        ...paragraph.attrs,
        tabs: [
          { val: "start", pos: 1500, leader: "underscore" },
          { val: "end", pos: 3000, leader: "hyphen" },
        ],
      },
    }),
  ).not.toBe(hash);
  expect(
    hashParagraphBlock({
      ...paragraph,
      attrs: { ...paragraph.attrs, tabs: paragraph.attrs.tabs.toReversed() },
    }),
  ).not.toBe(hash);
});

test("paragraph cache is order-independent across distinct tab stops", () => {
  withFakeTextMeasure(
    () => {
      const measureWidth = (position: number): number => {
        const measure = measureBlock(
          {
            kind: "paragraph",
            id: `tab-${position}`,
            runs: [
              { kind: "text", text: "A", fontSize: 11 },
              { kind: "tab" },
              { kind: "text", text: "B", fontSize: 11 },
            ],
            attrs: { tabs: [{ val: "start", pos: position }] },
          },
          400,
        );
        if (measure.kind !== "paragraph") {
          throw new Error("expected paragraph measure");
        }
        return measure.lines.at(0)?.width ?? 0;
      };

      expect([measureWidth(1500), measureWidth(3000)]).toEqual([105, 205]);
      clearAllCaches();
      expect([measureWidth(3000), measureWidth(1500)]).toEqual([205, 105]);
    },
    { charWidth: fixedCharWidth(5) },
  );
});

test("paragraph cache is order-independent across list paragraph mark sizes", () => {
  withFakeTextMeasure(
    () => {
      const measureLastLineHeight = (listParagraphMarkFontSize: number): number => {
        const measure = measureBlock(
          {
            kind: "paragraph",
            id: `list-paragraph-mark-${listParagraphMarkFontSize}`,
            runs: [{ kind: "text", text: "aaaa aaaa", fontFamily: "Aptos", fontSize: 10.5 }],
            attrs: {
              defaultFontSize: 10.5,
              defaultFontFamily: "Aptos",
              listMarker: "\u2022",
              listMarkerSuffix: "nothing",
              listParagraphMarkFontSize,
            },
          },
          60,
        );
        if (measure.kind !== "paragraph") {
          throw new Error("expected paragraph measure");
        }
        return measure.lines.at(-1)?.lineHeight ?? 0;
      };

      const smallThenLarge = [measureLastLineHeight(11), measureLastLineHeight(18)];
      expect(smallThenLarge[1]).toBeGreaterThan(smallThenLarge[0] ?? 0);
      clearAllCaches();
      const largeThenSmall = [measureLastLineHeight(18), measureLastLineHeight(11)];
      expect(largeThenSmall).toEqual(smallThenLarge.toReversed());
    },
    { charWidth: fixedCharWidth(10) },
  );
});

test("paragraph cache preserves line spacing units", () => {
  withFakeTextMeasure(() => {
    const paragraph = {
      kind: "paragraph",
      id: "spacing-unit",
      runs: [{ kind: "text", text: "Arabic", fontSize: 10 }],
      attrs: { spacing: { line: 2, lineRule: "auto", lineUnit: "multiplier" } },
    } as const satisfies ParagraphBlock;

    expectCachedMeasurementsMatchFreshInBothOrders(
      paragraph,
      {
        ...paragraph,
        attrs: { spacing: { ...paragraph.attrs.spacing, lineUnit: "px" } },
      },
      200,
    );
  });
});

test("paragraph cache preserves field, math, and rendered-break runs", () => {
  withFakeTextMeasure(() => {
    const field = {
      kind: "paragraph",
      id: "field",
      runs: [{ kind: "field", fieldType: "OTHER", fallback: "1", fontSize: 11 }],
    } as const satisfies ParagraphBlock;
    expectCachedMeasurementsMatchFreshInBothOrders(
      field,
      { ...field, runs: [{ ...field.runs[0], fallback: "123456789" }] },
      200,
    );

    const math = {
      kind: "paragraph",
      id: "math",
      runs: [
        {
          kind: "math",
          display: "inline",
          ommlXml: "<m:oMath/>",
          plainText: "x",
          fontSize: 11,
        },
      ],
    } as const satisfies ParagraphBlock;
    expectCachedMeasurementsMatchFreshInBothOrders(
      math,
      { ...math, runs: [{ ...math.runs[0], plainText: "x + y + z" }] },
      200,
    );

    const textRuns = [
      { kind: "text", text: "A", fontSize: 11 },
      { kind: "text", text: "B", fontSize: 11 },
    ] as const;
    const plain = {
      kind: "paragraph",
      id: "plain",
      runs: [...textRuns],
    } as const satisfies ParagraphBlock;
    const withRenderedBreak = {
      ...plain,
      runs: [textRuns[0], { kind: "renderedPageBreak" }, textRuns[1]],
    } as const satisfies ParagraphBlock;
    expectCachedMeasurementsMatchFreshInBothOrders(plain, withRenderedBreak, 200);
  });
});

test("paragraph cache preserves inferred RTL and tab typography", () => {
  withFakeTextMeasure(() => {
    const rtlTab = {
      kind: "paragraph",
      id: "rtl-tab",
      runs: [
        { kind: "text", text: "عنوان عربي", fontSize: 11, rtl: true },
        { kind: "tab", fontSize: 11 },
        { kind: "text", text: "1", fontSize: 11, rtl: true },
      ],
      attrs: {
        indent: { left: 48, right: 48, hanging: 48 },
        tabs: [{ val: "end", pos: 9000, leader: "dot" }],
      },
    } as const satisfies ParagraphBlock;
    expectCachedMeasurementsMatchFreshInBothOrders(
      rtlTab,
      {
        ...rtlTab,
        runs: [
          { ...rtlTab.runs[0], rtl: false },
          rtlTab.runs[1],
          { ...rtlTab.runs[2], rtl: false },
        ],
      },
      624,
    );

    const smallTab = {
      kind: "paragraph",
      id: "tab-typography",
      runs: [{ kind: "tab", fontFamily: "Aptos", fontSize: 10 }],
    } as const satisfies ParagraphBlock;
    expectCachedMeasurementsMatchFreshInBothOrders(
      smallTab,
      { ...smallTab, runs: [{ ...smallTab.runs[0], fontSize: 30 }] },
      200,
    );
  });
});

test("paragraph cache preserves image measurement geometry without source payloads", () => {
  withFakeTextMeasure(() => {
    const inlineImage = {
      kind: "paragraph",
      id: "inline-image",
      runs: [
        {
          kind: "image",
          src: "data:image/png;base64,large-payload-a",
          width: 40,
          height: 20,
        },
      ],
    } as const satisfies ParagraphBlock;
    const rotatedImage = {
      ...inlineImage,
      runs: [{ ...inlineImage.runs[0], transform: "rotate(90deg)" }],
    } as const satisfies ParagraphBlock;
    expectCachedMeasurementsMatchFreshInBothOrders(inlineImage, rotatedImage, 200);

    const compactBlockImage = {
      ...inlineImage,
      runs: [
        {
          ...inlineImage.runs[0],
          displayMode: "block",
          distTop: 6,
          distBottom: 6,
        },
      ],
    } as const satisfies ParagraphBlock;
    const spacedBlockImage = {
      ...compactBlockImage,
      runs: [{ ...compactBlockImage.runs[0], distTop: 20, distBottom: 20 }],
    } as const satisfies ParagraphBlock;
    expectCachedMeasurementsMatchFreshInBothOrders(compactBlockImage, spacedBlockImage, 200);

    expect(hashParagraphBlock(inlineImage)).toBe(
      hashParagraphBlock({
        ...inlineImage,
        runs: [{ ...inlineImage.runs[0], src: "data:image/png;base64,large-payload-b" }],
      }),
    );
  });
});
