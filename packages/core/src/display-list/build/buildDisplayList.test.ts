import { describe, expect, test } from "bun:test";

import { layoutDocument } from "../../layout-engine/index";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../../layout-engine/measure/measureBlocks";
import type {
  FlowBlock,
  ImageBlock,
  Layout,
  LayoutOptions,
  PageMargins,
  ParagraphBlock,
  TableBlock,
  TextBoxBlock,
} from "../../layout-engine/types";
import type { BlockLookup } from "../../layout-painter/index";
import type { DisplayGlyphRun, DisplayHitRegion, DisplayList, DisplayPrimitive } from "../types";
import { buildDisplayList } from "./buildDisplayList";
import { UNSUPPORTED_CONSTRUCT } from "./unsupported";

const CHAR_WIDTH_PX = 5;
const fakeMeasure = { charWidth: fixedCharWidth(CHAR_WIDTH_PX) };

const PAGE_SIZE = { w: 816, h: 1056 };
const MARGINS: PageMargins = { top: 96, right: 96, bottom: 96, left: 96 };
const CONTENT_WIDTH = PAGE_SIZE.w - MARGINS.left - MARGINS.right;

const LAYOUT_OPTIONS: LayoutOptions = { pageSize: PAGE_SIZE, margins: MARGINS };

type Built = { layout: Layout; blockLookup: BlockLookup };

const buildLayout = (blocks: FlowBlock[]): Built => {
  const measures = measureBlocks(blocks, CONTENT_WIDTH);
  const blockLookup: BlockLookup = new Map();
  for (const [index, block] of blocks.entries()) {
    const measure = measures[index];
    if (measure) {
      blockLookup.set(String(block.id), { block, measure });
    }
  }
  return { layout: layoutDocument(blocks, measures, LAYOUT_OPTIONS), blockLookup };
};

const buildFrom = (blocks: FlowBlock[]) => buildDisplayList(buildLayout(blocks));

const pagePrimitives = (blocks: FlowBlock[]): readonly DisplayPrimitive[] => {
  const list = buildFrom(blocks);
  const page = list.pages.at(0);
  expect(page).toBeDefined();
  // Index 0 is always the page background rect; the body starts after it.
  return page?.primitives.slice(1) ?? [];
};

const glyphRuns = (primitives: readonly DisplayPrimitive[]): DisplayGlyphRun[] =>
  primitives.filter((primitive): primitive is DisplayGlyphRun => primitive.kind === "glyphRun");

const flattenRegions = (regions: readonly DisplayHitRegion[]): DisplayHitRegion[] =>
  regions.flatMap((region) => [region, ...flattenRegions(region.children)]);

const para = (id: string, text: string, attrs?: ParagraphBlock["attrs"]): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [{ kind: "text", text }],
  ...(attrs === undefined ? {} : { attrs }),
});

describe("buildDisplayList: paragraph geometry", () => {
  test("a left-aligned line starts at the content-left edge, on the measured baseline", () => {
    withFakeTextMeasure(() => {
      const primitives = pagePrimitives([para("p", "Hello")]);
      const runs = glyphRuns(primitives);

      expect(runs).toHaveLength(1);
      const run = runs[0];
      expect(run?.text).toBe("Hello");
      expect(run?.xPx).toBe(MARGINS.left);
      expect(run?.advancesPx).toEqual([5, 5, 5, 5, 5]);

      // The line's own metrics decide the baseline: CSS centres the inline box
      // in the line box, so half the leading sits above the ascent.
      const { layout } = buildLayout([para("p", "Hello")]);
      const fragment = layout.pages[0]?.fragments[0];
      const line = measureBlocks([para("p", "Hello")], CONTENT_WIDTH)[0];
      expect(fragment?.kind).toBe("paragraph");
      if (fragment?.kind === "paragraph" && line?.kind === "paragraph") {
        const measured = line.lines[0];
        expect(measured).toBeDefined();
        if (measured) {
          const expected =
            fragment.y +
            (measured.lineHeight - (measured.ascent + measured.descent)) / 2 +
            measured.ascent;
          expect(run?.baselineYPx).toBeCloseTo(expected, 6);
        }
      }
    }, fakeMeasure);
  });

  test("right and centre alignment place the line by its measured width", () => {
    withFakeTextMeasure(() => {
      const text = "Hello";
      const widthPx = text.length * CHAR_WIDTH_PX;

      const right = glyphRuns(pagePrimitives([para("r", text, { alignment: "right" })]));
      expect(right[0]?.xPx).toBeCloseTo(MARGINS.left + CONTENT_WIDTH - widthPx, 6);

      const centre = glyphRuns(pagePrimitives([para("c", text, { alignment: "center" })]));
      expect(centre[0]?.xPx).toBeCloseTo(MARGINS.left + (CONTENT_WIDTH - widthPx) / 2, 6);
    }, fakeMeasure);
  });

  test("a hanging-indent list line places the marker before the body text", () => {
    withFakeTextMeasure(() => {
      const block = para("list", "Item text", {
        listMarker: "1.",
        listIsBullet: false,
        indent: { left: 48, hanging: 48 },
      });
      const runs = glyphRuns(pagePrimitives([block]));

      expect(runs).toHaveLength(2);
      const [marker, body] = runs;
      expect(marker?.text).toBe("1.");
      expect(body?.text).toBe("Item text");

      // Marker sits at `left - hanging`; the body picks up at the left indent
      // because a fitting marker's tab lands exactly there.
      expect(marker?.xPx).toBeCloseTo(MARGINS.left, 6);
      expect(body?.xPx).toBeCloseTo(MARGINS.left + 48, 6);
      expect(body?.xPx).toBeGreaterThan(marker?.xPx ?? 0);
    }, fakeMeasure);
  });
});

describe("buildDisplayList: tabs and justification", () => {
  test("a tab advances the cursor to its stop, measured from the fragment edge", () => {
    withFakeTextMeasure(() => {
      // 1440 twips = 96px from the fragment's own left edge, which is NOT the
      // page's: feeding the page-absolute x here would miss the stop entirely
      // and fall back to the default grid.
      const block: ParagraphBlock = {
        kind: "paragraph",
        id: "tab",
        runs: [{ kind: "text", text: "A" }, { kind: "tab" }, { kind: "text", text: "B" }],
        attrs: { tabs: [{ val: "start", pos: 1440 }] },
      };
      const runs = glyphRuns(pagePrimitives([block]));

      expect(runs.map((run) => run.text)).toEqual(["A", "B"]);
      expect(runs[0]?.xPx).toBe(MARGINS.left);
      expect(runs[1]?.xPx).toBeCloseTo(MARGINS.left + 96, 6);
    }, fakeMeasure);
  });

  test("tab and empty-line regions carry the positions a caret resolves", () => {
    withFakeTextMeasure(() => {
      const tabbed: ParagraphBlock = {
        kind: "paragraph",
        id: "tabbed",
        pmStart: 10,
        pmEnd: 14,
        runs: [
          { kind: "text", text: "A", pmStart: 11, pmEnd: 12 },
          { kind: "tab", pmStart: 12, pmEnd: 13 },
          { kind: "text", text: "B", pmStart: 13, pmEnd: 14 },
        ],
      };
      const blank: ParagraphBlock = {
        kind: "paragraph",
        id: "blank",
        pmStart: 20,
        pmEnd: 21,
        runs: [],
      };
      const regions = flattenRegions(buildFrom([tabbed, blank]).pages.at(0)?.regions ?? []);
      const tab = regions.find((region) => region.kind === "tab");
      const empty = regions.find(
        (region) => region.kind === "emptyRun" && region.model?.blockId === "blank",
      );

      expect(tab?.model?.pmRange).toMatchObject({ start: 12, end: 13 });
      expect(tab?.rect.widthPx).toBeGreaterThan(0);
      expect(empty?.model?.pmRange).toMatchObject({ start: 21, end: 21 });
    }, fakeMeasure);
  });

  test("a leader tab paints its fill characters inside a clip of the tab's box", () => {
    withFakeTextMeasure(() => {
      const block: ParagraphBlock = {
        kind: "paragraph",
        id: "toc",
        runs: [{ kind: "text", text: "Chapter" }, { kind: "tab" }, { kind: "text", text: "7" }],
        attrs: { tabs: [{ val: "start", pos: 1440, leader: "dot" }] },
      };
      const primitives = pagePrimitives([block]);
      const clip = primitives.find((primitive) => primitive.kind === "clipGroup");

      expect(clip?.kind).toBe("clipGroup");
      if (clip?.kind === "clipGroup") {
        const tabStartPx = MARGINS.left + "Chapter".length * CHAR_WIDTH_PX;
        expect(clip.rect.xPx).toBeCloseTo(tabStartPx, 6);
        expect(clip.rect.widthPx).toBeCloseTo(96 - "Chapter".length * CHAR_WIDTH_PX, 6);
        const leader = glyphRuns(clip.children).at(0);
        expect(leader?.text.startsWith(".")).toBe(true);
        expect(new Set(leader?.text ?? "")).toEqual(new Set("."));
      }
    }, fakeMeasure);
  });

  test("a justified non-final line spends its slack on the spaces, not on a backend", () => {
    withFakeTextMeasure(() => {
      const block = para("j", `${"word ".repeat(40).trim()}`, { alignment: "justify" });
      const { layout, blockLookup } = buildLayout([block]);
      const measure = blockLookup.get("j")?.measure;
      expect(measure?.kind).toBe("paragraph");
      if (measure?.kind !== "paragraph" || measure.lines.length < 2) {
        throw new Error("test needs a paragraph that wraps");
      }

      const list = buildDisplayList({ layout, blockLookup });
      const runs = glyphRuns(list.pages[0]?.primitives ?? []);
      const first = runs.at(0);
      expect(first).toBeDefined();
      if (!first) {
        return;
      }

      const paintedWidth = first.advancesPx.reduce((sum, advance) => sum + advance, 0);
      // The first line now fills the column exactly; every space carries an
      // equal share of the slack the measurer left.
      expect(paintedWidth).toBeCloseTo(CONTENT_WIDTH, 6);
      expect(first.xPx + paintedWidth).toBeCloseTo(MARGINS.left + CONTENT_WIDTH, 6);

      // A collapsed trailing space carries no advance, which is why the sum
      // above lands on the column width rather than one space past it.
      const characters = [...first.text];
      const interiorSpaces = characters.flatMap((char, index) =>
        char === " " && index < characters.length - 1 ? [first.advancesPx[index] ?? 0] : [],
      );
      expect(interiorSpaces.length).toBeGreaterThan(0);
      for (const advance of interiorSpaces) {
        expect(advance).toBeGreaterThan(CHAR_WIDTH_PX);
        expect(advance).toBeCloseTo(interiorSpaces[0] ?? 0, 6);
      }
      if (characters.at(-1) === " ") {
        expect(first.advancesPx.at(-1)).toBe(0);
      }
    }, fakeMeasure);
  });
});

describe("buildDisplayList: paint order", () => {
  test("a highlighted run emits its background rect before its glyphs", () => {
    withFakeTextMeasure(() => {
      const block: ParagraphBlock = {
        kind: "paragraph",
        id: "hl",
        runs: [{ kind: "text", text: "Marked", highlight: "#FFFF00" }],
      };
      const primitives = pagePrimitives([block]);

      const rectIndex = primitives.findIndex((primitive) => primitive.kind === "rect");
      const glyphIndex = primitives.findIndex((primitive) => primitive.kind === "glyphRun");
      expect(rectIndex).toBeGreaterThanOrEqual(0);
      expect(glyphIndex).toBeGreaterThan(rectIndex);

      const rect = primitives[rectIndex];
      expect(rect?.kind).toBe("rect");
      if (rect?.kind === "rect") {
        expect(rect.fill).toEqual({ r: 255, g: 255, b: 0, a: 1 });
        expect(rect.rect.widthPx).toBeCloseTo("Marked".length * CHAR_WIDTH_PX, 6);
      }
    }, fakeMeasure);
  });

  test("a table cell emits background, then borders, then its inner text", () => {
    withFakeTextMeasure(() => {
      const table: TableBlock = {
        kind: "table",
        id: "t",
        columnWidths: [200],
        rows: [
          {
            id: "r0",
            cells: [
              {
                id: "c0",
                background: "#EEEEEE",
                borders: {
                  top: { width: 1, style: "solid", color: "#333333" },
                  bottom: { width: 1, style: "solid", color: "#333333" },
                  left: { width: 1, style: "solid", color: "#333333" },
                  right: { width: 1, style: "solid", color: "#333333" },
                },
                blocks: [para("cp", "Cell")],
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
              },
            ],
          },
        ],
      };

      const primitives = pagePrimitives([table]);
      const kinds = primitives.map((primitive) => primitive.kind);

      const backgroundIndex = kinds.indexOf("rect");
      const firstLineIndex = kinds.indexOf("line");
      const glyphIndex = kinds.indexOf("glyphRun");

      expect(backgroundIndex).toBe(0);
      expect(firstLineIndex).toBeGreaterThan(backgroundIndex);
      expect(glyphIndex).toBeGreaterThan(firstLineIndex);
      // All four authored edges are painted on a single-cell table.
      expect(kinds.filter((kind) => kind === "line")).toHaveLength(4);
      expect(glyphRuns(primitives)[0]?.text).toBe("Cell");
    }, fakeMeasure);
  });

  test("image and text-box cell blocks keep nested hit regions", () => {
    withFakeTextMeasure(() => {
      const image: ImageBlock = {
        kind: "image",
        id: "cell-image",
        src: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
        width: 20,
        height: 10,
      };
      const textBox: TextBoxBlock = {
        kind: "textBox",
        id: "cell-box",
        width: 100,
        height: 30,
        content: [para("box-paragraph", "Inside")],
      };
      const table: TableBlock = {
        kind: "table",
        id: "t",
        columnWidths: [200],
        rows: [
          {
            id: "r0",
            cells: [
              {
                id: "c0",
                blocks: [image, textBox],
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
              },
            ],
          },
        ],
      };

      const regions = flattenRegions(buildFrom([table]).pages.at(0)?.regions ?? []);
      expect(
        regions
          .filter((region) => region.kind === "image" || region.kind === "textBox")
          .map((region) => [region.kind, region.model?.blockId]),
      ).toEqual([
        ["image", "cell-image"],
        ["textBox", "cell-box"],
      ]);
      const box = regions.find((region) => region.model?.blockId === "cell-box");
      expect(
        flattenRegions(box?.children ?? []).some((region) => region.kind === "paragraph"),
      ).toBe(true);
    }, fakeMeasure);
  });

  test("applies DrawingML rotation and flips about its authored text-box frame centre", () => {
    withFakeTextMeasure(() => {
      const textBox: TextBoxBlock = {
        kind: "textBox",
        id: "rotated-box",
        width: 120,
        height: 40,
        transform: "rotate(270deg) scaleX(-1) scaleY(-1)",
        content: [para("rotated-box-paragraph", "Likelihood")],
      };

      const primitive = pagePrimitives([textBox]).at(0);
      expect(primitive?.kind).toBe("rotateGroup");
      if (primitive?.kind !== "rotateGroup") {
        return;
      }
      expect(primitive.degrees).toBe(270);
      expect(primitive.originXPx).toBe(MARGINS.left + 60);
      expect(primitive.originYPx).toBe(MARGINS.top + 20);
      expect(primitive.scaleX).toBe(-1);
      expect(primitive.scaleY).toBe(-1);
      expect(primitive.children.some(({ kind }) => kind === "glyphRun")).toBe(true);
    }, fakeMeasure);
  });
});

describe("buildDisplayList: contract obligations", () => {
  test("two builds of the same layout are deeply equal", () => {
    withFakeTextMeasure(() => {
      const blocks = (): FlowBlock[] => [
        para("a", "Alpha", { alignment: "center" }),
        para("b", "Beta", { listMarker: "•", indent: { left: 36, hanging: 36 } }),
      ];
      const built = buildLayout(blocks());
      expect(buildDisplayList(built)).toEqual(buildDisplayList(built));
    }, fakeMeasure);
  });

  test("a hidden run is reported as unsupported and paints nothing", () => {
    withFakeTextMeasure(() => {
      const block: ParagraphBlock = {
        kind: "paragraph",
        id: "h",
        runs: [{ kind: "text", text: "Secret", hidden: true }],
      };
      const list = buildFrom([block]);
      const page = list.pages.at(0);

      expect(glyphRuns(page?.primitives ?? [])).toHaveLength(0);
      const reported = list.unsupported.filter(
        (entry) => entry.construct === UNSUPPORTED_CONSTRUCT.hiddenRun,
      );
      expect(reported).toHaveLength(1);
      expect(reported[0]?.pageIndex).toBe(0);
      expect(reported[0]?.detail).toContain("w:vanish");
    }, fakeMeasure);
  });

  test("black run colour is re-materialized, because the DOM painter drops it", () => {
    withFakeTextMeasure(() => {
      const block: ParagraphBlock = {
        kind: "paragraph",
        id: "k",
        runs: [{ kind: "text", text: "Ink", color: "#000000" }],
      };
      const runs = glyphRuns(pagePrimitives([block]));
      expect(runs[0]?.color).toEqual({ r: 0, g: 0, b: 0, a: 1 });
    }, fakeMeasure);
  });

  test("a cell block with no measure is named rather than dropped", () => {
    withFakeTextMeasure(() => {
      const table: TableBlock = {
        kind: "table",
        id: "t",
        columnWidths: [200],
        rows: [
          {
            id: "r0",
            cells: [
              {
                id: "c0",
                blocks: [para("cp", "Cell")],
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
              },
            ],
          },
        ],
      };
      const built = buildLayout([table]);
      const measure = built.blockLookup.get("t")?.measure;
      expect(measure?.kind).toBe("table");
      if (measure?.kind === "table") {
        // The pair the builder walks: one block, no measure beside it.
        measure.rows[0]?.cells[0]?.blocks.pop();
      }

      const list = buildDisplayList(built);
      expect(glyphRuns(list.pages[0]?.primitives ?? [])).toHaveLength(0);
      const reported = list.unsupported.filter(
        (entry) => entry.construct === UNSUPPORTED_CONSTRUCT.missingBlock,
      );
      expect(reported).toHaveLength(1);
      expect(reported[0]?.detail).toContain("cp");
      expect(reported[0]?.detail).toContain("measure");
    }, fakeMeasure);
  });

  test("the font table interns one face per distinct family/weight/slant", () => {
    withFakeTextMeasure(() => {
      const block: ParagraphBlock = {
        kind: "paragraph",
        id: "f",
        runs: [
          { kind: "text", text: "plain " },
          { kind: "text", text: "bold ", bold: true },
          { kind: "text", text: "plain again" },
        ],
      };
      const list = buildFrom([block]);
      expect(list.fonts).toHaveLength(2);
      expect(list.fonts.map((face) => face.weight)).toEqual([400, 700]);
      expect(glyphRuns(list.pages[0]?.primitives ?? []).map((run) => run.font)).toEqual([0, 1, 0]);
    }, fakeMeasure);
  });
});

describe("buildDisplayList: layout-invisible features", () => {
  const constructsOf = (list: DisplayList): string[] =>
    list.unsupported.map((entry) => entry.construct);

  test("a document with neither page borders nor a watermark reports neither", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({
        ...buildLayout([para("a", "Alpha")]),
        documentFeatures: { pageBorders: false, watermark: false },
      });
      expect(list.unsupported).toEqual([]);
    }, fakeMeasure);
  });

  test("only the feature the document has is reported", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({
        ...buildLayout([para("a", "Alpha")]),
        documentFeatures: { pageBorders: true, watermark: false },
      });
      expect(constructsOf(list)).toEqual([UNSUPPORTED_CONSTRUCT.pageBorders]);
      expect(list.unsupported.at(0)?.pageIndex).toBe(0);
    }, fakeMeasure);
  });

  test("an unstated document reports both, because unknown is not absent", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList(buildLayout([para("a", "Alpha")]));
      expect(constructsOf(list)).toEqual([
        UNSUPPORTED_CONSTRUCT.pageBorders,
        UNSUPPORTED_CONSTRUCT.watermark,
      ]);
    }, fakeMeasure);
  });

  test("a layout with no page reports no gap, because there is no page to name", () => {
    const list = buildDisplayList({
      layout: { pageSize: PAGE_SIZE, pages: [] },
      blockLookup: new Map(),
    });
    expect(list.pages).toEqual([]);
    expect(list.unsupported).toEqual([]);
  });
});
