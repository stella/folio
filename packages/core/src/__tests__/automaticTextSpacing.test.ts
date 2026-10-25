import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { buildGlyphs } from "../display-list/build/glyphs";
import { serializeParagraph } from "../docx/serializer/paragraphSerializer";
import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import {
  clickToPositionInParagraph,
  positionToX,
} from "../layout-bridge/engine/clickToPosition";
import {
  getCaretPosition,
  selectionToRects,
} from "../layout-bridge/engine/selectionRects";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../layout-engine/measure/__tests__/fakeTextMeasure";
import {
  automaticRunGap,
  automaticTextGap,
  automaticTextSpacingGaps,
} from "../layout-engine/measure/automaticTextSpacing";
import { measureTableBlock } from "../layout-engine/measure/measureBlocks";
import { buildRunFontStyle } from "../layout-engine/measure/measureHelpers";
import { measureParagraph } from "../layout-engine/measure/measureParagraph";
import {
  getMeasureProvider,
  measurePositionedRun,
  measureRun,
  measureText,
  measureTextWidth,
  setMeasureProvider,
} from "../layout-engine/measure/measureProvider";
import type { FontStyle } from "../layout-engine/measure/measureTypes";
import { canonicalLineWidth } from "../layout-engine/measure/paragraphMeasureShared";
import type {
  Layout,
  ParagraphBlock,
  ParagraphFragment,
  RunFormatting,
  TableBlock,
  TextRun,
} from "../layout-engine/types";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { ParagraphDirection } from "../prosemirror/paragraphDirection";
import { schema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";

const style = {
  fontFamily: "Arial",
  fontSize: 12,
  automaticTextSpacing: { latin: true, numbers: true },
} satisfies FontStyle;

setDefaultTimeout(propertyTestTimeout(10_000));

const paragraph = (runs: TextRun[]): ParagraphBlock => ({
  kind: "paragraph",
  id: "automatic-spacing",
  runs,
});

const literal = (text: string): TextRun => ({ ...style, kind: "text", text });

test("paragraphs without inter-script candidates keep their existing flow shape", () => {
  for (const text of ["ordinary Latin text", "123", "漢字", "漢α"]) {
    const block = toFlowBlocks(
      schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text(text)]),
      ]),
    ).at(0);
    expect(block?.kind).toBe("paragraph");
    if (block?.kind !== "paragraph")
      throw new TypeError("Expected paragraph flow.");
    expect(block.runs.at(0)).not.toHaveProperty("automaticTextSpacing");
  }
});

test("unstyled geometry uses the same default em as paragraph measurement", () => {
  withFakeTextMeasure(
    () => {
      const block = {
        kind: "paragraph",
        id: "unstyled-spacing",
        pmStart: 0,
        pmEnd: 4,
        runs: [
          {
            kind: "text",
            text: "漢A",
            pmStart: 1,
            pmEnd: 3,
            automaticTextSpacing: { latin: true, numbers: true },
          },
        ],
      } satisfies ParagraphBlock;
      const measured = measureParagraph(block, 1000);
      const position = positionToX(block, measured, 3, 1000);
      expect(position).not.toBeNull();
      expect(canonicalLineWidth(position?.x ?? 0)).toBe(
        measured.lines.at(0)?.width,
      );
    },
    { charWidth: fixedCharWidth(10) },
  );
});

test("literal policy follows canonical manual and automatic paragraph direction", () => {
  const cases = [
    { direction: null, enabled: true },
    { direction: { source: "auto" }, enabled: false },
    { direction: { source: "manual", value: "rtl" }, enabled: false },
    { direction: { source: "manual", value: "ltr" }, enabled: true },
  ] as const satisfies readonly {
    direction: ParagraphDirection | null;
    enabled: boolean;
  }[];
  for (const entry of cases) {
    const block = toFlowBlocks(
      schema.node("doc", null, [
        schema.node("paragraph", { direction: entry.direction }, [
          schema.text("漢A"),
        ]),
      ]),
    ).at(0);
    expect(block?.kind).toBe("paragraph");
    if (block?.kind !== "paragraph")
      throw new TypeError("Expected paragraph flow.");
    const run = block.runs.at(0);
    const policy = run?.kind === "text" ? run.automaticTextSpacing : undefined;
    expect(policy !== undefined).toBe(entry.enabled);
  }
});

test("literal source is explicit across content controls, symbols, and projected field children", () => {
  const field = schema.node(
    "structuredField",
    {
      fieldType: "REF",
      instruction: "REF target",
      displayText: "漢A",
      fieldKind: "simple",
    },
    [schema.text("漢A"), schema.node("pageBreakRun"), schema.text("漢1")],
  );
  const node = schema.node("paragraph", null, [
    schema.text("漢A"),
    schema.node("sdt", null, [schema.text("漢1")]),
    schema.node("symbol", { font: "Arial", char: "6F22" }),
    field,
  ]);
  const runs = toFlowBlocks(schema.node("doc", null, [node])).flatMap(
    (block) => (block.kind === "paragraph" ? block.runs : []),
  );
  expect(
    runs.map((run) =>
      "automaticTextSpacing" in run ? run.automaticTextSpacing : undefined,
    ),
  ).toEqual([
    { latin: true, numbers: true },
    { latin: true, numbers: true },
    undefined,
    undefined,
    undefined,
  ]);
});

test("shaping normalizes only glyph widths, never the automatic advance", () => {
  withFakeTextMeasure(
    () => {
      const native = getMeasureProvider();
      setMeasureProvider({
        ...native,
        measureTextWidth: (text, font) =>
          native.measureTextWidth(text, font) - (text === "AV漢" ? 4 : 0),
      });
      const unspaced = measurePositionedRun({
        text: "AV漢",
        style: {
          ...style,
          automaticTextSpacing: { latin: false, numbers: false },
        },
      });
      const spaced = measurePositionedRun({ text: "AV漢", style });
      const gap = automaticTextSpacingGaps("AV漢", style).at(0)?.advancePx ?? 0;
      expect(spaced.charWidths.at(0)).toBeCloseTo(
        unspaced.charWidths.at(0) ?? 0,
        8,
      );
      expect(
        (spaced.charWidths.at(1) ?? 0) - (unspaced.charWidths.at(1) ?? 0),
      ).toBeCloseTo(gap, 8);
      expect(spaced.charWidths.at(2)).toBeCloseTo(
        unspaced.charWidths.at(2) ?? 0,
        8,
      );
      expect(
        spaced.charWidths.reduce((sum, width) => sum + width, 0),
      ).toBeCloseTo(spaced.width, 8);
    },
    { charWidth: fixedCharWidth(10) },
  );
});

test("glyph advances, engine caret, selection and hit testing own identical source boundaries", () => {
  withFakeTextMeasure(
    () => {
      const text = "𠀀Aé\u0301漢1";
      const expected = measurePositionedRun({ text, style });
      for (const parts of [[text], ["𠀀", "Aé\u0301漢1"]]) {
        const document = createEmptyDocument();
        const first = document.package.document.content.at(0);
        if (first?.type !== "paragraph")
          throw new TypeError("Expected a paragraph.");
        first.content = parts.map((part, index) => ({
          type: "run",
          formatting: {
            fontFamily: { ascii: "Arial", hAnsi: "Arial" },
            fontSize: 24,
            color: { rgb: index === 0 ? "000000" : "112233" },
          },
          content: [{ type: "text", text: part }],
        }));
        const block = toFlowBlocks(toProseDoc(document)).at(0);
        if (block?.kind !== "paragraph")
          throw new TypeError("Expected paragraph flow.");
        const measured = measureParagraph(block, 1000);
        const fragment = {
          kind: "paragraph",
          blockId: block.id,
          x: 0,
          y: 0,
          width: 1000,
          height: measured.totalHeight,
          fromLine: 0,
          toLine: measured.lines.length,
        } satisfies ParagraphFragment;
        const layout = {
          pageGap: 0,
          pages: [
            {
              number: 1,
              size: { w: 1000, h: 1000 },
              margins: { top: 0, right: 0, bottom: 0, left: 0 },
              fragments: [fragment],
            },
          ],
        } satisfies Layout;
        for (const offset of [2, 5, 6, 7]) {
          const width = expected.charWidths
            .slice(0, offset)
            .reduce((sum, advance) => sum + advance, 0);
          expect(positionToX(block, measured, offset + 1, 1000)?.x).toBeCloseTo(
            width,
            8,
          );
          expect(
            getCaretPosition(layout, [block], [measured], offset + 1)?.x,
          ).toBeCloseTo(width, 8);
          expect(
            selectionToRects(layout, [block], [measured], 1, offset + 1).at(0)
              ?.width,
          ).toBeCloseTo(width, 8);
          expect(
            clickToPositionInParagraph({
              fragment,
              block,
              measure: measured,
              pageIndex: 0,
              localY: measured.totalHeight / 2,
              localX: width + 0.01,
            })?.pmPosition,
          ).toBe(offset + 1);
        }
      }
    },
    { charWidth: fixedCharWidth(10) },
  );
});

test("formatting-only splits preserve unwrapped width and display advances", () => {
  withFakeTextMeasure(
    () => {
      assertProperty(
        fc.property(
          fc.array(
            fc.constantFrom("漢", "𠀀", "A", "é\u0301", "1", " ", "🙂"),
            { minLength: 1, maxLength: 40 },
          ),
          (parts) => {
            const runs = parts.map(literal);
            const expected = measureTextWidth(parts.join(""), style);
            const visibleWidth = measureTextWidth(
              parts.join("").trimEnd(),
              style,
            );
            const actual = runs.reduce(
              (width, run, index) =>
                width +
                measureTextWidth(run.text, style) +
                automaticRunGap(run, runs[index + 1]),
              0,
            );
            expect(actual).toBeCloseTo(expected, 8);
            expect(
              measureParagraph(paragraph(runs), 100000).lines.reduce(
                (width, line) => width + line.width,
                0,
              ),
            ).toBeCloseTo(visibleWidth, 8);
            const glyphWidth = runs.reduce(
              (width, run, index) =>
                width +
                buildGlyphs({
                  text: run.text,
                  style,
                  allCaps: false,
                  collapsed: false,
                  spaceDeltaPx: 0,
                  trailingAutomaticSpacingPx: automaticRunGap(
                    run,
                    runs[index + 1],
                  ),
                }).widthPx,
              0,
            );
            expect(glyphWidth).toBeCloseTo(expected, 8);
          },
        ),
        { id: "split-run-advance-invariance", numRuns: 80 },
      );
    },
    { charWidth: fixedCharWidth(10) },
  );
});

test("a boundary advance is kept on one line and discarded after wrapping", () => {
  withFakeTextMeasure(
    () => {
      for (const runs of [[literal("漢A")], [literal("漢"), literal("A")]]) {
        expect(
          measureParagraph(paragraph(runs), 30).lines.map((line) => line.width),
        ).toEqual([24]);
        expect(
          measureParagraph(paragraph(runs), 20).lines.map((line) => line.width),
        ).toEqual([10, 10]);
      }
    },
    { charWidth: fixedCharWidth(10) },
  );
});

test("flow derives default and authored switches without changing the source model", () => {
  withFakeTextMeasure(
    () => {
      for (const enabled of [undefined, false, true]) {
        const document = createEmptyDocument({ initialText: "漢A漢1" });
        const first = document.package.document.content.at(0);
        if (first?.type !== "paragraph")
          throw new TypeError("Expected a paragraph.");
        first.formatting = {
          ...first.formatting,
          ...(enabled === undefined
            ? {}
            : { autoSpaceDE: enabled, autoSpaceDN: enabled }),
        };
        const before = JSON.stringify(document);
        const xmlBefore = serializeParagraph(first);
        const pmDoc = toProseDoc(document);
        const pmBefore = pmDoc.toJSON();
        const block = toFlowBlocks(pmDoc).at(0);
        expect(block?.kind).toBe("paragraph");
        if (block?.kind !== "paragraph")
          throw new TypeError("Expected paragraph flow.");
        expect(block.runs.at(0)).toMatchObject({
          automaticTextSpacing: {
            latin: enabled !== false,
            numbers: enabled !== false,
          },
        });
        measureParagraph(block, 100);
        expect(JSON.stringify(document)).toBe(before);
        expect(serializeParagraph(first)).toBe(xmlBefore);
        expect(pmDoc.toJSON()).toEqual(pmBefore);
      }
    },
    { charWidth: fixedCharWidth(10) },
  );
});

test("derived advances exclude projected, note, script and direction overrides", () => {
  const exclusions = [
    { rtl: true },
    { hidden: true },
    { superscript: true },
    { subscript: true },
    { positionPx: 1 },
    { footnoteRefId: 1 },
    { endnoteRefId: 1 },
    { templatePreview: "plain" },
  ] as const satisfies readonly (RunFormatting & {
    templatePreview?: TextRun["templatePreview"];
  })[];
  for (const override of exclusions)
    expect(
      automaticTextSpacingGaps(
        "漢A",
        buildRunFontStyle({ ...literal("漢A"), ...override }, "Arial", 12),
      ),
    ).toEqual([]);
});

test("auto-width noWrap cells price unwrapped advances without changing the authored grid", () => {
  withFakeTextMeasure(
    () => {
      assertProperty(
        fc.property(
          fc.integer({ min: 1, max: 3 }),
          fc.integer({ min: 1, max: 12 }),
          (span, repeats) => {
            const text = "漢A".repeat(repeats);
            const block = paragraph([literal(text)]);
            const table = {
              kind: "table",
              id: "automatic-table",
              layout: "autofit",
              columnWidths: Array.from({ length: span }, () => 1),
              rows: [
                {
                  id: "automatic-row",
                  cells: [
                    {
                      id: "automatic-cell",
                      colSpan: span,
                      noWrap: true,
                      padding: { left: 0, right: 0, top: 0, bottom: 0 },
                      blocks: [block],
                    },
                  ],
                },
              ],
            } satisfies TableBlock;
            const before = JSON.stringify(table);
            const measured = measureTableBlock(table, 1);
            const content = measured.rows.at(0)?.cells.at(0)?.blocks.at(0);
            expect(content?.kind).toBe("paragraph");
            if (content?.kind !== "paragraph")
              throw new TypeError("Expected measured paragraph.");
            expect(content.lines.at(0)?.width).toBeCloseTo(
              measureTextWidth(text, style),
              8,
            );
            expect(measured.columnWidths).toEqual(table.columnWidths);
            expect(JSON.stringify(table)).toBe(before);
            expect(
              measureTableBlock({ ...table, layout: "fixed" }, 1).columnWidths,
            ).toEqual(table.columnWidths);
          },
        ),
        { id: "no-wrap-cell-advance-invariance", numRuns: 24 },
      );
    },
    { charWidth: fixedCharWidth(10) },
  );
});

describe("automatic inter-script spacing", () => {
  test("keeps width and UTF-16 advances consistent without a trailing gap", () => {
    withFakeTextMeasure(
      () => {
        expect(measureTextWidth("漢A", style)).toBe(24);
        expect(measureText("漢A", style).width).toBe(24);
        expect(measureRun("漢A", style)).toMatchObject({
          width: 24,
          charWidths: [14, 10],
        });
        expect(measureRun("𠀀A", style)).toMatchObject({
          width: 24,
          charWidths: [14, 0, 10],
        });
        expect(measureTextWidth("漢", style)).toBe(10);
        expect(measureTextWidth("A", style)).toBe(10);
      },
      { charWidth: fixedCharWidth(10) },
    );
  });

  test("uses the East Asian em before Latin letters and digits", () => {
    for (const text of ["漢A", "漢1", "漢é", "あA", "한1"]) {
      expect(automaticTextSpacingGaps(text, style)).toEqual([
        { offset: 1, advancePx: 4 },
      ]);
    }
  });

  test("uses the authored Latin face average at a reverse boundary", () => {
    expect(automaticTextSpacingGaps("A漢", style)).toEqual([
      { offset: 1, advancePx: (16 * 904) / 2048 / 2 },
    ]);
    expect(
      automaticTextSpacingGaps("A漢", { ...style, fontFamily: "Unknown Face" }),
    ).toEqual([]);
    expect(
      automaticTextSpacingGaps(
        "A漢",
        buildRunFontStyle(
          { automaticTextSpacing: style.automaticTextSpacing },
          "Arial",
          12,
        ),
      ),
    ).toEqual([]);
  });

  test("independently honors Latin and numeric switches", () => {
    expect(
      automaticTextSpacingGaps("漢A漢1", {
        ...style,
        automaticTextSpacing: { latin: false, numbers: true },
      }),
    ).toEqual([{ offset: 3, advancePx: 4 }]);
    expect(
      automaticTextSpacingGaps("漢1漢A", {
        ...style,
        automaticTextSpacing: { latin: true, numbers: false },
      }),
    ).toEqual([{ offset: 3, advancePx: 4 }]);
  });

  test("keeps source offsets and combining sequences intact", () => {
    expect(automaticTextSpacingGaps("😀a\u0301漢A", style)).toEqual([
      { offset: 4, advancePx: (16 * 904) / 2048 / 2 },
      { offset: 5, advancePx: 4 },
    ]);
  });

  test("does not stack automatic spacing on manual or nonliteral advances", () => {
    for (const text of [
      "漢 A",
      "漢\tA",
      "漢Ａ",
      "漢α",
      "漢١",
      "漢😀",
      "漢。",
      "漢漢",
    ]) {
      expect(automaticTextSpacingGaps(text, style)).toEqual([]);
    }
    for (const override of [
      { letterSpacing: 1 },
      { letterSpacing: -1 },
      { horizontalScale: 120 },
      { forceComplexScript: true },
      { fontVariant: "small-caps" },
      { textTransform: "uppercase" },
    ] as const) {
      expect(
        automaticTextGap({
          before: { character: "漢", style: { ...style, ...override } },
          after: { character: "A", style },
          policy: style.automaticTextSpacing,
        }),
      ).toBe(0);
    }
  });
});
