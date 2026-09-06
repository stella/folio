/**
 * The half of paint equivalence that needs neither a browser nor a rasterizer.
 *
 * `scripts/paint-equivalence.ts` measures how far apart the two backends paint
 * one display list; that needs `mutool` and a chromium, so it cannot gate a
 * merge. What can be asserted here is the part a raster score would only
 * report indirectly: that both backends cover every primitive kind, that they
 * agree on the page geometry, and that the producer and the PDF writer are
 * deterministic. A page can score 1.0 because neither backend painted
 * something, which is exactly why coverage is asserted separately from
 * similarity.
 *
 * This file is under `src` on purpose: `bun test src` is the package's test
 * command, so a test in `scripts/` would never run. Being under `src` it may
 * not import from `scripts/`, and does not.
 */

import { inflateSync } from "node:zlib";

import { describe, expect, test } from "bun:test";

import { layoutDocument } from "../layout-engine/index";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../layout-engine/measure/measureBlocks";
import type { FlowBlock, LayoutOptions, PageMargins } from "../layout-engine/types";
import type { BlockLookup } from "../layout-painter/index";
import { POINTS_PER_PIXEL } from "../pdf/pageSpace";
import type { PdfFontSource } from "../pdf/writePdf";
import { writePdf } from "../pdf/writePdf";
import { buildDisplayList } from "./build/buildDisplayList";
import { renderDisplayListToDom } from "./dom/renderDisplayListToDom";
import { BLACK, DISPLAY_PRIMITIVE_KINDS, WHITE } from "./primitives";
import type {
  DisplayFontFace,
  DisplayImageSource,
  DisplayList,
  DisplayPage,
  DisplayPrimitive,
} from "./types";

const TIMESTAMP = "2026-01-02T03:04:05Z";
const PAGE_WIDTH_PX = 816;
const PAGE_HEIGHT_PX = 1056;
const SECOND_PAGE_WIDTH_PX = 1056;
const SECOND_PAGE_HEIGHT_PX = 816;

// ---------------------------------------------------------------------------
// A document just wide enough for the DOM backend
// ---------------------------------------------------------------------------

/**
 * Bun's test runtime has no DOM and the repository carries neither happy-dom
 * nor jsdom, so the backend is driven against the narrow surface it uses.
 */
type StubElement = {
  readonly tagName: string;
  readonly style: Record<string, string>;
  readonly dataset: Record<string, string>;
  readonly attributes: Record<string, string>;
  readonly children: StubElement[];
  className: string;
  id: string;
  textContent: string;
  href: string;
  title: string;
  src: string;
  alt: string;
  readonly append: (...nodes: StubElement[]) => void;
  readonly setAttribute: (name: string, value: string) => void;
};

const createStubElement = (tagName: string): StubElement => {
  const children: StubElement[] = [];
  const attributes: Record<string, string> = {};
  return {
    tagName,
    style: {},
    dataset: {},
    attributes,
    children,
    className: "",
    id: "",
    textContent: "",
    href: "",
    title: "",
    src: "",
    alt: "",
    append: (...nodes) => {
      children.push(...nodes);
    },
    setAttribute: (name, value) => {
      attributes[name] = value;
    },
  };
};

// SAFETY: the backend calls `createElement` and then touches only the
// properties StubElement declares; nothing here reaches the rest of the DOM.
const stubDocument = () => ({ createElement: createStubElement }) as unknown as Document;

// SAFETY: every element compared here came from `stubDocument()`.
const asStub = (element: HTMLElement) => element as unknown as StubElement;

const renderToStubs = (list: DisplayList): StubElement[] =>
  renderDisplayListToDom(list, { doc: stubDocument() }).map(asStub);

// ---------------------------------------------------------------------------
// Reading the PDF back
// ---------------------------------------------------------------------------

const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);

/**
 * Every text stream in the file, decompressed. An image or font-program
 * stream names a payload that is not text, so it is skipped rather than
 * inflated into noise.
 */
const textStreams = (bytes: Uint8Array): readonly string[] => {
  const text = latin1(bytes);
  const out: string[] = [];
  // The dictionary capture must not run back through an earlier `<<`, or it
  // picks up keys belonging to the object before this one.
  const pattern = /<<((?:(?!<<)[^])*?)\/Length (\d+)>>\nstream\n/g;
  let match = pattern.exec(text);
  while (match !== null) {
    const [, dict = "", length = "0"] = match;
    const start = match.index + match[0].length;
    if (dict.includes("/FlateDecode") && !dict.includes("/Subtype") && !dict.includes("/Length1")) {
      out.push(latin1(inflateSync(bytes.subarray(start, start + Number(length)))));
    }
    match = pattern.exec(text);
  }
  return out;
};

/** Page content streams, told apart from other text by the base matrix each
 * one opens with, in page order. */
const contentStreams = (bytes: Uint8Array): readonly string[] =>
  textStreams(bytes).filter((stream) => stream.includes(" cm\n"));

type PageBoxPt = { readonly widthPt: number; readonly heightPt: number };

const mediaBoxes = (bytes: Uint8Array): readonly PageBoxPt[] => {
  const pattern = /\/MediaBox \[0 0 ([\d.]+) ([\d.]+)\]/g;
  return [...latin1(bytes).matchAll(pattern)].map(([, width = "0", height = "0"]) => ({
    widthPt: Number(width),
    heightPt: Number(height),
  }));
};

// ---------------------------------------------------------------------------
// One display list, both backends
// ---------------------------------------------------------------------------

const FACE: DisplayFontFace = {
  family: "Helvetica",
  weight: 400,
  italic: false,
  generic: "sans-serif",
  fontBoxAscentRatio: 0.9,
  fontBoxDescentRatio: 0.2,
};

/** A real 1x1 PNG: the PDF backend decodes image bytes, so a signature stub
 * would fail the writer rather than exercise the primitive. */
const ONE_PIXEL_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const IMAGE: DisplayImageSource = {
  format: "png",
  bytes: new Uint8Array(Buffer.from(ONE_PIXEL_PNG_BASE64, "base64")),
  pixelWidth: 1,
  pixelHeight: 1,
};

const NO_FONTS: PdfFontSource = { load: () => [] };

/**
 * One primitive per discriminator. `satisfies Record<...>` is what makes a new
 * kind a compile error here rather than a mark one backend quietly never
 * paints. The groups carry no children on purpose: a group with a child would
 * pass this test on the child's marks alone.
 */
const PRIMITIVE_SAMPLES = {
  glyphRun: {
    kind: "glyphRun",
    font: 0,
    fontSizePx: 16,
    color: BLACK,
    xPx: 72,
    baselineYPx: 100,
    text: "abc",
    advancesPx: [9, 9, 9],
    direction: "ltr",
  },
  rect: {
    kind: "rect",
    rect: { xPx: 10, yPx: 20, widthPx: 100, heightPx: 30 },
    fill: BLACK,
    stroke: { color: BLACK, thicknessPx: 1, pattern: "solid" },
  },
  line: {
    kind: "line",
    x1Px: 10,
    y1Px: 200,
    x2Px: 300,
    y2Px: 200,
    stroke: { color: BLACK, thicknessPx: 2, pattern: "dashed" },
  },
  image: {
    kind: "image",
    image: 0,
    rect: { xPx: 0, yPx: 0, widthPx: 100, heightPx: 100 },
    opacity: 1,
  },
  clipGroup: {
    kind: "clipGroup",
    rect: { xPx: 0, yPx: 0, widthPx: 10, heightPx: 10 },
    children: [],
  },
  rotateGroup: {
    kind: "rotateGroup",
    degrees: 45,
    originXPx: 5,
    originYPx: 5,
    children: [],
  },
  opacityGroup: {
    kind: "opacityGroup",
    opacity: 0.5,
    children: [],
  },
} as const satisfies Record<DisplayPrimitive["kind"], DisplayPrimitive>;

const pageWith = (primitives: readonly DisplayPrimitive[]): DisplayPage => ({
  pageNumber: 1,
  widthPx: PAGE_WIDTH_PX,
  heightPx: PAGE_HEIGHT_PX,
  orientation: "portrait",
  primitives,
  links: [],
});

const listOf = (pages: readonly DisplayPage[]): DisplayList => ({
  pages,
  fonts: [FACE],
  images: [IMAGE],
  outline: [],
  metadata: {},
  unsupported: [],
});

const write = (list: DisplayList) => {
  const result = writePdf(list, { fonts: NO_FONTS, timestamp: TIMESTAMP });
  if (result.isErr()) {
    throw result.error;
  }
  return result.value;
};

describe("both backends cover every primitive kind", () => {
  const emptyPageStreamLength = contentStreams(write(listOf([pageWith([])])).bytes).at(0)?.length;

  test("the empty page establishes a floor for the coverage comparison", () => {
    expect(emptyPageStreamLength).toBeGreaterThan(0);
  });

  for (const kind of DISPLAY_PRIMITIVE_KINDS) {
    test(`${kind} reaches both backends`, () => {
      const list = listOf([pageWith([PRIMITIVE_SAMPLES[kind]])]);

      const domPage = renderToStubs(list).at(0);
      expect(
        domPage?.children.length ?? 0,
        `the DOM backend painted nothing for ${kind}`,
      ).toBeGreaterThan(0);

      const stream = contentStreams(write(list).bytes).at(0) ?? "";
      expect(
        stream.length,
        `the PDF backend emitted no content stream for ${kind}`,
      ).toBeGreaterThan(0);
      expect(stream.length, `the PDF backend added no marks for ${kind}`).toBeGreaterThan(
        emptyPageStreamLength ?? 0,
      );
    });
  }
});

describe("one display list, two backends", () => {
  const list = listOf([
    pageWith([PRIMITIVE_SAMPLES.rect, PRIMITIVE_SAMPLES.glyphRun]),
    {
      pageNumber: 2,
      widthPx: SECOND_PAGE_WIDTH_PX,
      heightPx: SECOND_PAGE_HEIGHT_PX,
      orientation: "landscape",
      primitives: [PRIMITIVE_SAMPLES.line],
      links: [],
    },
  ]);

  test("both backends emit one page per display page", () => {
    expect(renderToStubs(list)).toHaveLength(list.pages.length);
    expect(mediaBoxes(write(list).bytes)).toHaveLength(list.pages.length);
  });

  test("both backends give each page the box the display list states", () => {
    const domPages = renderToStubs(list);
    const boxes = mediaBoxes(write(list).bytes);

    for (const [index, page] of list.pages.entries()) {
      const dom = domPages.at(index);
      expect(dom?.style["width"]).toBe(`${page.widthPx}px`);
      expect(dom?.style["height"]).toBe(`${page.heightPx}px`);
      // PDF's unit is the point, so the same box arrives scaled, not different.
      expect(boxes.at(index)?.widthPt).toBeCloseTo(page.widthPx * POINTS_PER_PIXEL, 6);
      expect(boxes.at(index)?.heightPt).toBeCloseTo(page.heightPx * POINTS_PER_PIXEL, 6);
    }
  });
});

describe("determinism", () => {
  test("writePdf twice over one list and timestamp is byte-identical", () => {
    const list = listOf([pageWith(DISPLAY_PRIMITIVE_KINDS.map((kind) => PRIMITIVE_SAMPLES[kind]))]);

    const first = write(list).bytes;
    const second = write(list).bytes;

    expect(first.length).toBe(second.length);
    expect(Buffer.from(first).equals(Buffer.from(second))).toBe(true);
  });

  const CHAR_WIDTH_PX = 5;
  const PAGE_SIZE = { w: PAGE_WIDTH_PX, h: PAGE_HEIGHT_PX };
  const MARGINS: PageMargins = { top: 96, right: 96, bottom: 96, left: 96 };
  const LAYOUT_OPTIONS: LayoutOptions = { pageSize: PAGE_SIZE, margins: MARGINS };

  const paragraph = (id: string, text: string): FlowBlock => ({
    kind: "paragraph",
    id,
    runs: [{ kind: "text", text }],
  });

  test("buildDisplayList twice over one layout is deeply equal", () => {
    withFakeTextMeasure(
      () => {
        const blocks = [paragraph("a", "Hello display list"), paragraph("b", "Second paragraph")];
        const measures = measureBlocks(blocks, PAGE_SIZE.w - MARGINS.left - MARGINS.right);
        const blockLookup: BlockLookup = new Map(
          blocks.flatMap((block, index) => {
            const measure = measures.at(index);
            return measure === undefined ? [] : [[String(block.id), { block, measure }] as const];
          }),
        );
        const layout = layoutDocument(blocks, measures, LAYOUT_OPTIONS);

        const first = buildDisplayList({ layout, blockLookup, pageBackground: WHITE });
        const second = buildDisplayList({ layout, blockLookup, pageBackground: WHITE });

        expect(first).toEqual(second);
      },
      { charWidth: fixedCharWidth(CHAR_WIDTH_PX) },
    );
  });
});
