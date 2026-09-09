/**
 * The four constructs that reach the builder as inputs rather than through
 * `Layout`, plus the two facts a run carries for an editing surface.
 *
 * Each construct is asserted three ways, because the three cases must not
 * collapse into one: supplied means painted, present-but-withheld means
 * reported, and absent from the document means neither.
 */

import { describe, expect, test } from "bun:test";

import { layoutDocument } from "../../layout-engine/index";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../../layout-engine/measure/measureBlocks";
import type {
  FlowBlock,
  FootnoteContent,
  HeaderFooterContent,
  ImageBlock,
  Layout,
  LayoutOptions,
  PageMargins,
  ParagraphBlock,
  TextBoxBlock,
} from "../../layout-engine/types";
import type { BlockLookup } from "../../layout-painter/index";
import type { EmbeddedFont } from "../../fonts/embeddedFonts";
import type { DisplayGlyphRun, DisplayHitRegion, DisplayList, DisplayPrimitive } from "../types";
import { buildDisplayList, type BuildDisplayListOptions } from "./buildDisplayList";
import { UNSUPPORTED_CONSTRUCT } from "./unsupported";

const CHAR_WIDTH_PX = 5;
const fakeMeasure = { charWidth: fixedCharWidth(CHAR_WIDTH_PX) };

const PAGE_SIZE = { w: 816, h: 1056 };
const MARGINS: PageMargins = { top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 };
const CONTENT_WIDTH = PAGE_SIZE.w - MARGINS.left - MARGINS.right;

const para = (id: string, text: string, attrs?: ParagraphBlock["attrs"]): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [{ kind: "text", text }],
  ...(attrs === undefined ? {} : { attrs }),
});

type Built = { layout: Layout; blockLookup: BlockLookup };

const buildLayout = (blocks: FlowBlock[], options: Partial<LayoutOptions> = {}): Built => {
  const measures = measureBlocks(blocks, CONTENT_WIDTH);
  const blockLookup: BlockLookup = new Map();
  for (const [index, block] of blocks.entries()) {
    const measure = measures[index];
    if (measure) {
      blockLookup.set(String(block.id), { block, measure });
    }
  }
  const layoutOptions: LayoutOptions = { pageSize: PAGE_SIZE, margins: MARGINS, ...options };
  return { layout: layoutDocument(blocks, measures, layoutOptions), blockLookup };
};

const storyContent = (id: string, text: string): HeaderFooterContent => {
  const blocks = [para(id, text)];
  const measures = measureBlocks(blocks, CONTENT_WIDTH);
  const measure = measures.at(0);
  return {
    blocks,
    measures,
    height: measure?.kind === "paragraph" ? measure.totalHeight : 0,
  };
};

const storyBlocks = (blocks: FlowBlock[]): HeaderFooterContent => {
  const measures = measureBlocks(blocks, CONTENT_WIDTH);
  let height = 0;
  for (const measure of measures) {
    switch (measure.kind) {
      case "paragraph":
      case "table":
        height += measure.totalHeight;
        break;
      case "image":
      case "textBox":
        height += measure.height;
        break;
      case "sectionBreak":
      case "pageBreak":
      case "columnBreak":
        break;
      default:
        measure satisfies never;
    }
  }
  return {
    blocks,
    measures,
    height,
  };
};

const footnoteContent = (id: number, text: string): FootnoteContent => {
  const blocks = [para(`fn-${String(id)}`, text)];
  const measures = measureBlocks(blocks, CONTENT_WIDTH);
  const measure = measures.at(0);
  return {
    id,
    displayNumber: id,
    blocks,
    measures,
    height: measure?.kind === "paragraph" ? measure.totalHeight : 0,
  };
};

const flatten = (primitives: readonly DisplayPrimitive[]): DisplayPrimitive[] =>
  primitives.flatMap((primitive) =>
    primitive.kind === "clipGroup" ||
    primitive.kind === "rotateGroup" ||
    primitive.kind === "opacityGroup"
      ? [primitive, ...flatten(primitive.children)]
      : [primitive],
  );

const glyphRuns = (list: DisplayList, pageIndex = 0): DisplayGlyphRun[] =>
  flatten(list.pages.at(pageIndex)?.primitives ?? []).filter(
    (primitive): primitive is DisplayGlyphRun => primitive.kind === "glyphRun",
  );

const constructsOf = (list: DisplayList): string[] =>
  list.unsupported.map((entry) => entry.construct);

const flattenRegions = (regions: readonly DisplayHitRegion[]): DisplayHitRegion[] =>
  regions.flatMap((region) => [region, ...flattenRegions(region.children)]);

const BORDERS: NonNullable<BuildDisplayListOptions["pageBorders"]> = {
  top: { style: "single", size: 8, space: 24, color: { rgb: "FF0000" } },
  bottom: { style: "single", size: 8, space: 24, color: { rgb: "FF0000" } },
  left: { style: "single", size: 8, space: 24, color: { rgb: "FF0000" } },
  right: { style: "single", size: 8, space: 24, color: { rgb: "FF0000" } },
  offsetFrom: "page",
  zOrder: "back",
};

describe("page borders", () => {
  test("supplied borders paint one stroke per authored side", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({ ...buildLayout([para("a", "A")]), pageBorders: BORDERS });
      const lines = (list.pages.at(0)?.primitives ?? []).filter(
        (primitive) => primitive.kind === "line",
      );

      expect(lines).toHaveLength(4);
      expect(constructsOf(list)).not.toContain(UNSUPPORTED_CONSTRUCT.pageBorders);
      // `offsetFrom: "page"` insets each side by `w:space`; a 1pt stroke is
      // centred on its path, so the line sits half a thickness further in.
      const top = lines.find((line) => line.kind === "line" && line.y1Px === line.y2Px);
      expect(top?.kind).toBe("line");
      if (top?.kind === "line") {
        expect(top.y1Px).toBeCloseTo(32 + top.stroke.thicknessPx / 2, 6);
        expect(top.stroke.color).toEqual({ r: 255, g: 0, b: 0, a: 1 });
      }
    }, fakeMeasure);
  });

  test("a zOrder=back border paints under the body, a front border over it", () => {
    withFakeTextMeasure(() => {
      const back = buildDisplayList({ ...buildLayout([para("a", "A")]), pageBorders: BORDERS });
      const front = buildDisplayList({
        ...buildLayout([para("a", "A")]),
        pageBorders: { ...BORDERS, zOrder: "front" },
      });

      const firstGlyph = (list: DisplayList): number =>
        (list.pages.at(0)?.primitives ?? []).findIndex(
          (primitive) => primitive.kind === "glyphRun",
        );
      const firstLine = (list: DisplayList): number =>
        (list.pages.at(0)?.primitives ?? []).findIndex((primitive) => primitive.kind === "line");

      expect(firstLine(back)).toBeLessThan(firstGlyph(back));
      expect(firstLine(front)).toBeGreaterThan(firstGlyph(front));
    }, fakeMeasure);
  });

  test("a document that has borders the caller withheld is reported", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({
        ...buildLayout([para("a", "A")]),
        documentFeatures: { pageBorders: true, watermark: false },
      });
      expect(constructsOf(list)).toEqual([UNSUPPORTED_CONSTRUCT.pageBorders]);
    }, fakeMeasure);
  });
});

describe("watermark", () => {
  test("a text watermark paints one rotated, translucent run centred on the page", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({
        ...buildLayout([para("a", "A")]),
        watermark: { kind: "text", text: "DRAFT" },
      });
      const primitives = list.pages.at(0)?.primitives ?? [];
      const rotate = primitives.find((primitive) => primitive.kind === "rotateGroup");

      expect(rotate?.kind).toBe("rotateGroup");
      if (rotate?.kind !== "rotateGroup") {
        return;
      }
      expect(rotate.degrees).toBe(-45);
      expect(rotate.originXPx).toBeCloseTo(PAGE_SIZE.w / 2, 6);
      expect(rotate.originYPx).toBeCloseTo(PAGE_SIZE.h / 2, 6);

      const opacity = rotate.children.at(0);
      expect(opacity?.kind).toBe("opacityGroup");
      if (opacity?.kind !== "opacityGroup") {
        return;
      }
      expect(opacity.opacity).toBe(0.5);
      const run = opacity.children.at(0);
      expect(run?.kind).toBe("glyphRun");
      if (run?.kind === "glyphRun") {
        expect(run.text).toBe("DRAFT");
        const widthPx = run.advancesPx.reduce((sum, advance) => sum + advance, 0);
        expect(run.xPx + widthPx / 2).toBeCloseTo(PAGE_SIZE.w / 2, 6);
        // No model position: a watermark lives in a header shape, not in the
        // body the caller edits.
        expect(run.pmRange).toBeUndefined();
      }
      expect(constructsOf(list)).not.toContain(UNSUPPORTED_CONSTRUCT.watermark);
    }, fakeMeasure);
  });

  test("a picture watermark with no resolved source is reported, not dropped", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({
        ...buildLayout([para("a", "A")]),
        watermark: { kind: "picture", imageRId: "rId7" },
        documentFeatures: { pageBorders: false, watermark: true },
      });
      expect(constructsOf(list)).toEqual([UNSUPPORTED_CONSTRUCT.watermark]);
      expect(list.unsupported.at(0)?.detail).toContain("rId7");
    }, fakeMeasure);
  });

  test("a picture watermark uses its authored shape box", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({
        ...buildLayout([para("a", "A")]),
        watermark: {
          kind: "picture",
          imageRId: "rId7",
          widthPt: 300,
          heightPt: 120,
        },
        watermarkImageSrc:
          "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
      });
      const image = (list.pages.at(0)?.primitives ?? []).find(
        (primitive) => primitive.kind === "image",
      );

      expect(image?.kind).toBe("image");
      if (image?.kind !== "image") {
        return;
      }
      expect(image.rect).toEqual({
        xPx: 208,
        yPx: 448,
        widthPx: 400,
        heightPx: 160,
      });
      expect(image.opacity).toBe(0.18);
    }, fakeMeasure);
  });

  test("a document that has a watermark the caller withheld is reported", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({
        ...buildLayout([para("a", "A")]),
        documentFeatures: { pageBorders: false, watermark: true },
      });
      expect(constructsOf(list)).toEqual([UNSUPPORTED_CONSTRUCT.watermark]);
    }, fakeMeasure);
  });
});

const FOOTNOTE_ID = 7;

const footnoteBearingLayout = (): Built =>
  buildLayout(
    [
      {
        kind: "paragraph",
        id: "body",
        runs: [
          { kind: "text", text: "Body" },
          { kind: "text", text: "1", footnoteRefId: FOOTNOTE_ID, superscript: true },
        ],
      },
    ],
    { footnoteHeightById: new Map([[FOOTNOTE_ID, 20]]) },
  );

describe("footnote bodies", () => {
  test("a supplied body paints inside the reserved band, under the separator rule", () => {
    withFakeTextMeasure(() => {
      const built = footnoteBearingLayout();
      const page = built.layout.pages.at(0);
      expect(page?.footnoteIds).toEqual([FOOTNOTE_ID]);

      const list = buildDisplayList({
        ...built,
        footnoteContentById: new Map([[FOOTNOTE_ID, footnoteContent(FOOTNOTE_ID, "Note text")]]),
      });
      const note = glyphRuns(list).find((run) => run.text === "Note text");
      const bandTopPx =
        MARGINS.top +
        (PAGE_SIZE.h - MARGINS.top - MARGINS.bottom) -
        (page?.footnoteReservedHeight ?? 0);

      expect(note).toBeDefined();
      expect(note?.baselineYPx ?? 0).toBeGreaterThan(bandTopPx);
      expect(note?.baselineYPx ?? 0).toBeLessThan(PAGE_SIZE.h - MARGINS.bottom);
      expect(note?.xPx).toBe(MARGINS.left);
      // The band still opens with its rule.
      const rule = (list.pages.at(0)?.primitives ?? []).find(
        (primitive) =>
          primitive.kind === "rect" &&
          primitive.rect.widthPx < CONTENT_WIDTH &&
          primitive.rect.yPx > bandTopPx - 1,
      );
      expect(rule?.kind).toBe("rect");
      expect(constructsOf(list)).not.toContain(UNSUPPORTED_CONSTRUCT.footnoteContent);
    }, fakeMeasure);
  });

  test("a page whose bodies were withheld reports them and paints only the rule", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList(footnoteBearingLayout());
      expect(constructsOf(list)).toContain(UNSUPPORTED_CONSTRUCT.footnoteContent);
      expect(list.unsupported.at(-1)?.detail).toContain(String(FOOTNOTE_ID));
      expect(glyphRuns(list).map((run) => run.text)).toEqual(["Body", "1"]);
    }, fakeMeasure);
  });

  test("a footnote run carries no model range: the body's positions are not its own", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({
        ...footnoteBearingLayout(),
        footnoteContentById: new Map([[FOOTNOTE_ID, footnoteContent(FOOTNOTE_ID, "Note text")]]),
      });
      expect(glyphRuns(list).find((run) => run.text === "Note text")?.pmRange).toBeUndefined();
    }, fakeMeasure);
  });
});

const SECTION_REFS = [{ headerDefault: "rIdH", footerDefault: "rIdF" }];

describe("header and footer stories", () => {
  test("a supplied header paints its text through the shared paragraph builder", () => {
    withFakeTextMeasure(() => {
      const built = buildLayout([para("a", "Body")], {
        sectionHeaderFooterRefs: SECTION_REFS,
      });
      const list = buildDisplayList({
        ...built,
        headerContentByRId: new Map([["rIdH", storyContent("h", "Header")]]),
        footerContentByRId: new Map([["rIdF", storyContent("f", "Footer")]]),
      });
      const runs = glyphRuns(list);
      const header = runs.find((run) => run.text === "Header");
      const footer = runs.find((run) => run.text === "Footer");
      const body = runs.find((run) => run.text === "Body");

      expect(header).toBeDefined();
      expect(footer).toBeDefined();
      // Same builder as the body: same left edge, same per-code-point advances.
      expect(header?.xPx).toBe(MARGINS.left);
      expect(header?.advancesPx).toEqual(Array.from<number>({ length: 6 }).fill(CHAR_WIDTH_PX));
      expect(header?.font).toBe(body?.font);
      // The header sits above the body, the footer below it.
      expect(header?.baselineYPx ?? 0).toBeLessThan(body?.baselineYPx ?? 0);
      expect(footer?.baselineYPx ?? 0).toBeGreaterThan(body?.baselineYPx ?? 0);
      expect(footer?.baselineYPx ?? 0).toBeGreaterThan(PAGE_SIZE.h - MARGINS.bottom);
      expect(constructsOf(list)).not.toContain(UNSUPPORTED_CONSTRUCT.headerFooterContent);
    }, fakeMeasure);
  });

  test("a header run carries no model range: a header is its own story", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({
        ...buildLayout([para("a", "Body")], { sectionHeaderFooterRefs: SECTION_REFS }),
        headerContentByRId: new Map([["rIdH", storyContent("h", "Header")]]),
      });
      expect(glyphRuns(list).find((run) => run.text === "Header")?.pmRange).toBeUndefined();
    }, fakeMeasure);
  });

  test("header images and text boxes identify their blocks inside the header story", () => {
    withFakeTextMeasure(() => {
      const image: ImageBlock = {
        kind: "image",
        id: "header-image",
        src: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
        width: 20,
        height: 10,
      };
      const textBox: TextBoxBlock = {
        kind: "textBox",
        id: "header-box",
        width: 100,
        height: 30,
        content: [para("header-box-paragraph", "Inside")],
      };
      const list = buildDisplayList({
        ...buildLayout([para("a", "Body")], { sectionHeaderFooterRefs: SECTION_REFS }),
        headerContentByRId: new Map([["rIdH", storyBlocks([image, textBox])]]),
      });
      const regions = flattenRegions(list.pages.at(0)?.regions ?? []);

      expect(
        regions
          .filter((region) => region.kind === "image" || region.kind === "textBox")
          .map((region) => [region.kind, region.model?.blockId]),
      ).toEqual([
        ["image", "header-image"],
        ["textBox", "header-box"],
      ]);
      const box = regions.find((region) => region.model?.blockId === "header-box");
      expect(
        flattenRegions(box?.children ?? []).some((region) => region.kind === "paragraph"),
      ).toBe(true);
    }, fakeMeasure);
  });

  test("positions floating images from their own header and footer paragraph anchors", () => {
    withFakeTextMeasure(() => {
      const floatingStory = (id: string): HeaderFooterContent => {
        const blocks: FlowBlock[] = [
          {
            kind: "paragraph",
            id,
            runs: [
              {
                kind: "image",
                src: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
                width: 120,
                height: 20,
                wrapType: "square",
                position: {
                  horizontal: { relativeTo: "column", posOffset: 0 },
                  vertical: { relativeTo: "paragraph", posOffset: 91_440 },
                },
              },
            ],
          },
        ];
        return storyBlocks(blocks);
      };
      const header = floatingStory("header-float");
      const footer = floatingStory("footer-float");
      const list = buildDisplayList({
        ...buildLayout([para("body", "Body")], { sectionHeaderFooterRefs: SECTION_REFS }),
        headerContentByRId: new Map([["rIdH", header]]),
        footerContentByRId: new Map([["rIdF", footer]]),
      });
      const images = (list.pages.at(0)?.primitives ?? []).filter(
        (primitive) => primitive.kind === "image",
      );

      expect(images).toHaveLength(2);
      expect(images.at(0)?.kind).toBe("image");
      expect(images.at(1)?.kind).toBe("image");
      if (images.at(0)?.kind !== "image" || images.at(1)?.kind !== "image") {
        return;
      }
      expect(images.at(0)?.rect).toEqual({
        xPx: MARGINS.left,
        yPx: (MARGINS.header ?? 48) + 10,
        widthPx: 120,
        heightPx: 20,
      });
      expect(images.at(1)?.rect).toEqual({
        xPx: MARGINS.left,
        yPx: PAGE_SIZE.h - (MARGINS.footer ?? 48) - footer.height + 10,
        widthPx: 120,
        heightPx: 20,
      });
      expect(constructsOf(list)).not.toContain(UNSUPPORTED_CONSTRUCT.headerFooterContent);
    }, fakeMeasure);
  });

  test("a page that names header parts nobody supplied reports them", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList(
        buildLayout([para("a", "Body")], { sectionHeaderFooterRefs: SECTION_REFS }),
      );
      expect(constructsOf(list)).toContain(UNSUPPORTED_CONSTRUCT.headerFooterContent);
      expect(glyphRuns(list).map((run) => run.text)).toEqual(["Body"]);
    }, fakeMeasure);
  });

  test("a supplied map that lacks the part this page selects reports that page", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({
        ...buildLayout([para("a", "Body")], { sectionHeaderFooterRefs: SECTION_REFS }),
        headerContentByRId: new Map([["rIdH", storyContent("h", "Header")]]),
      });
      const footerGaps = list.unsupported.filter((entry) => entry.detail.includes("footer"));
      expect(footerGaps).toHaveLength(1);
    }, fakeMeasure);
  });
});

describe("a document with none of the four", () => {
  test("paints no furniture and reports none", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({
        ...buildLayout([para("a", "Body")]),
        documentFeatures: { pageBorders: false, watermark: false },
      });

      expect(list.unsupported).toEqual([]);
      // Background rect, then the one body run: nothing else.
      expect((list.pages.at(0)?.primitives ?? []).map((primitive) => primitive.kind)).toEqual([
        "rect",
        "glyphRun",
      ]);
    }, fakeMeasure);
  });
});

describe("determinism", () => {
  test("two builds of one layout and one set of furniture are deeply equal", () => {
    withFakeTextMeasure(() => {
      const built = footnoteBearingLayout();
      const options = {
        ...built,
        pageBorders: BORDERS,
        watermark: { kind: "text", text: "DRAFT" },
        headerContentByRId: new Map([["rIdH", storyContent("h", "Header")]]),
        footnoteContentById: new Map([[FOOTNOTE_ID, footnoteContent(FOOTNOTE_ID, "Note")]]),
      } as const satisfies BuildDisplayListOptions;

      expect(buildDisplayList(options)).toEqual(buildDisplayList(options));
    }, fakeMeasure);
  });
});

describe("embedded faces", () => {
  const face = (family: string, bytes: Uint8Array): EmbeddedFont => ({
    family: `folio-embedded-test-${family}`,
    originalFamily: family,
    style: "normal",
    weight: 400,
    bytes: bytes as Uint8Array<ArrayBuffer>,
    subsetted: false,
  });

  test("a face the measurer used carries its bytes once", () => {
    withFakeTextMeasure(() => {
      const bytes = new Uint8Array([0x00, 0x01, 0x00, 0x00]);
      const block: ParagraphBlock = {
        kind: "paragraph",
        id: "p",
        runs: [
          { kind: "text", text: "one", fontFamily: "Embedded Face" },
          { kind: "text", text: "two", fontFamily: "Embedded Face" },
        ],
      };
      const list = buildDisplayList({
        ...buildLayout([block]),
        embeddedFonts: [face("Embedded Face", bytes)],
      });

      const embedded = list.fonts.flatMap((entry) =>
        entry.embedded === undefined ? [] : [entry.embedded],
      );
      expect(embedded).toHaveLength(1);
      expect(embedded.at(0)?.bytes).toBe(bytes);
      expect(new Set(embedded.map((entry) => entry.id)).size).toBe(1);
    }, fakeMeasure);
  });

  test("a face the document never uses contributes nothing", () => {
    withFakeTextMeasure(() => {
      const list = buildDisplayList({
        ...buildLayout([para("a", "Body")]),
        embeddedFonts: [face("Unused Face", new Uint8Array([0x00]))],
      });
      expect(list.fonts.every((entry) => entry.embedded === undefined)).toBe(true);
    }, fakeMeasure);
  });
});

describe("model ranges", () => {
  test("a body run carries the range its text occupies, a list marker carries none", () => {
    withFakeTextMeasure(() => {
      const block: ParagraphBlock = {
        kind: "paragraph",
        id: "p",
        runs: [{ kind: "text", text: "Item", pmStart: 3, pmEnd: 7 }],
        attrs: { listMarker: "1.", indent: { left: 48, hanging: 48 } },
      };
      const runs = glyphRuns(buildDisplayList(buildLayout([block])));

      const marker = runs.find((run) => run.text === "1.");
      const body = runs.find((run) => run.text === "Item");
      expect(marker?.pmRange).toBeUndefined();
      expect(body?.pmRange).toEqual({ start: 3, end: 7, story: { kind: "body" } });
    }, fakeMeasure);
  });

  test("a wrapped paragraph gives each line the range of the characters on it", () => {
    withFakeTextMeasure(() => {
      const text = "word ".repeat(40).trim();
      const block: ParagraphBlock = {
        kind: "paragraph",
        id: "p",
        runs: [{ kind: "text", text, pmStart: 1, pmEnd: 1 + text.length }],
      };
      const runs = glyphRuns(buildDisplayList(buildLayout([block])));
      expect(runs.length).toBeGreaterThan(1);

      // Every line's range is the slice of the paragraph on that line, in
      // order: a later line never starts before an earlier one.
      let previousStart = 0;
      for (const run of runs) {
        expect(run.pmRange).toBeDefined();
        if (!run.pmRange) {
          continue;
        }
        expect(run.pmRange.start).toBeGreaterThanOrEqual(previousStart);
        expect(run.pmRange.end).toBeGreaterThan(run.pmRange.start);
        expect(run.pmRange.end).toBeLessThanOrEqual(1 + text.length);
        previousStart = run.pmRange.start;
      }
      expect(runs.at(0)?.pmRange?.start).toBe(1);
      expect(runs.at(-1)?.pmRange?.end).toBe(1 + text.length);
    }, fakeMeasure);
  });

  test("a substituted field value carries none: the model holds the field, not the digits", () => {
    withFakeTextMeasure(() => {
      const block: ParagraphBlock = {
        kind: "paragraph",
        id: "p",
        runs: [{ kind: "field", fieldType: "PAGE", instruction: "PAGE", pmStart: 2, pmEnd: 3 }],
      };
      const runs = glyphRuns(buildDisplayList(buildLayout([block])));
      expect(runs.at(0)?.text).toBe("1");
      expect(runs.at(0)?.pmRange).toBeUndefined();
    }, fakeMeasure);
  });
});
