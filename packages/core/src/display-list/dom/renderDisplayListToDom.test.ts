import { describe, expect, test } from "bun:test";

import { DISPLAY_PRIMITIVE_KINDS } from "../primitives";
import type {
  DisplayFontFace,
  DisplayGlyphRun,
  DisplayImageSource,
  DisplayList,
  DisplayPage,
  DisplayPrimitive,
  DisplayStrokePattern,
} from "../types";
import { renderDisplayListToDom, renderDisplayPageToDom } from "./renderDisplayListToDom";

/**
 * Bun's test runtime has no DOM and the repository carries neither happy-dom
 * nor jsdom, so the renderer is driven against the narrow surface it actually
 * uses. `style` and `dataset` are plain objects, which means a stub property
 * is camelCase where a browser would expose `data-advance-sum`.
 */
type StubElement = {
  readonly tagName: string;
  readonly style: Record<string, string>;
  readonly dataset: Record<string, string>;
  readonly children: StubElement[];
  readonly attributes: Record<string, string>;
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
    children,
    attributes,
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

// SAFETY: the renderer touches only createElement plus the element properties
// StubElement declares; nothing here reaches the rest of the DOM API.
const stubDocument = () => ({ createElement: createStubElement }) as unknown as Document;

const asStub = (element: HTMLElement) => element as unknown as StubElement;

const FONT: DisplayFontFace = {
  family: "Times New Roman",
  weight: 400,
  italic: false,
  generic: "serif",
  fontBoxAscentRatio: 0.75,
  fontBoxDescentRatio: 0.25,
};

const EMBEDDED_FONT: DisplayFontFace = {
  ...FONT,
  family: "Folio Sans",
  generic: "sans-serif",
  embedded: { id: "embedded-1", bytes: new Uint8Array([0, 1, 0, 0]) },
};

const IMAGE: DisplayImageSource = {
  format: "png",
  bytes: new Uint8Array([137, 80, 78, 71]),
  pixelWidth: 2,
  pixelHeight: 2,
};

const BLUE = { r: 0, g: 0, b: 255, a: 1 };

const emptyPage = (primitives: readonly DisplayPrimitive[]): DisplayPage => ({
  pageNumber: 1,
  widthPx: 816,
  heightPx: 1056,
  orientation: "portrait",
  primitives,
  links: [],
});

const renderPrimitives = (primitives: readonly DisplayPrimitive[]) =>
  asStub(
    renderDisplayPageToDom(emptyPage(primitives), {
      doc: stubDocument(),
      fonts: [FONT],
      images: [IMAGE],
      pageIndex: 0,
    }),
  );

/** The run's text as a document-order reader (a `Range`, a copy) sees it. */
const runText = (span: StubElement | undefined) => span?.textContent ?? "";

/**
 * One primitive per discriminator. `satisfies Record<...>` is what makes a new
 * kind a compile error here rather than a mark that quietly never renders.
 */
const PRIMITIVE_SAMPLES = {
  glyphRun: {
    kind: "glyphRun",
    font: 0,
    fontSizePx: 12,
    color: { r: 0, g: 0, b: 0, a: 1 },
    xPx: 100,
    baselineYPx: 50,
    text: "abc",
    advancesPx: [5, 6, 7],
    direction: "ltr",
  },
  rect: {
    kind: "rect",
    rect: { xPx: 10, yPx: 20, widthPx: 30, heightPx: 40 },
    fill: BLUE,
  },
  line: {
    kind: "line",
    x1Px: 0,
    y1Px: 10,
    x2Px: 100,
    y2Px: 10,
    stroke: { color: BLUE, thicknessPx: 1, pattern: "solid" },
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
} satisfies Record<DisplayPrimitive["kind"], DisplayPrimitive>;

const renderRun = (run: Partial<DisplayGlyphRun>) =>
  renderPrimitives([{ ...PRIMITIVE_SAMPLES.glyphRun, ...run }]).children.at(0);

const listOf = (
  pages: readonly DisplayPage[],
  fonts: readonly DisplayFontFace[] = [FONT],
): DisplayList => ({
  pages,
  fonts,
  images: [IMAGE],
  outline: [],
  metadata: {},
  unsupported: [],
});

describe("renderDisplayListToDom", () => {
  test("renders one page element per display page, in page order", () => {
    const list = listOf([emptyPage([]), { ...emptyPage([]), pageNumber: 7 }]);

    const pages = renderDisplayListToDom(list, { doc: stubDocument() }).map(asStub);

    expect(pages).toHaveLength(2);
    expect(pages.map((page) => page.className)).toEqual(["layout-page", "layout-page"]);
    // Anchors follow the list index, not the engine's physical page number.
    expect(pages.map((page) => page.id)).toEqual(["page-0", "page-1"]);
    expect(pages.at(0)?.style.width).toBe("816px");
  });

  test("paints every primitive kind", () => {
    for (const kind of DISPLAY_PRIMITIVE_KINDS) {
      const page = renderPrimitives([PRIMITIVE_SAMPLES[kind]]);
      expect(page.children.length, `no element painted for ${kind}`).toBeGreaterThan(0);
    }
  });

  test("places a glyph run on its own baseline, sized by its advances", () => {
    const span = renderPrimitives([PRIMITIVE_SAMPLES.glyphRun]).children.at(0);

    expect(span?.tagName).toBe("span");
    expect(span?.style.left).toBe("100px");
    // A 12px face whose font box is 0.75 above the baseline: the box top is
    // 9px above it and the line box is exactly the font box tall.
    expect(span?.style.top).toBe("41px");
    expect(span?.style.lineHeight).toBe("12px");
    expect(span?.style.width).toBe("18px");
    expect(span?.dataset.advanceSum).toBe("18");
    expect(runText(span)).toBe("abc");
    expect(span?.style.fontFamily).toBe(`"Times New Roman", serif`);
    expect(span?.style.direction).toBe("ltr");
    // Nothing may re-derive a position from inline flow, or re-order a run.
    expect(span?.style.textAlign).toBeUndefined();
    expect(span?.style.textIndent).toBeUndefined();
    expect(span?.style.unicodeBidi).toBeUndefined();
  });

  test("gives an rtl run the same box as an ltr one", () => {
    const ltr = renderPrimitives([PRIMITIVE_SAMPLES.glyphRun]).children.at(0);
    const rtl = renderPrimitives([{ ...PRIMITIVE_SAMPLES.glyphRun, direction: "rtl" }]).children.at(
      0,
    );

    expect(rtl?.style.left).toBe(ltr?.style.left);
    expect(rtl?.style.width).toBe(ltr?.style.width);
    expect(rtl?.style.direction).toBe("rtl");
  });

  test("emits a run's text as one shaped element", () => {
    // The display list places the origin; the browser shapes inside the run,
    // which is the only thing that can form a ligature, kern a pair, or pick a
    // cursive letter's positional form. Splitting the run into a box per code
    // point would override all three.
    const span = renderRun({ text: "abc", advancesPx: [5, 6, 7] });

    expect(runText(span)).toBe("abc");
    expect(span?.children.length ?? 0).toBe(0);
    expect(span?.style.width).toBe("18px");
    expect(span?.dataset.advanceSum).toBe("18");
  });

  test("keeps a cursively joined word whole", () => {
    // Nothing may split it: isolated forms are what a reader would see.
    const span = renderRun({ text: "\u0628\u064a\u062a", advancesPx: [6, 6, 6] });

    expect(runText(span)).toBe("\u0628\u064a\u062a");
    expect(span?.children.length ?? 0).toBe(0);
  });

  test("counts code points, not UTF-16 units, when sizing a run", () => {
    const span = renderRun({ text: "a\u{1f600}b", advancesPx: [5, 12, 5] });

    expect(span?.dataset.advanceSum).toBe("22");
    expect(runText(span)).toBe("a\u{1f600}b");
  });

  test("refuses a run whose advances do not match its code points", () => {
    expect(() => renderRun({ text: "abc", advancesPx: [5, 6] })).toThrow();
  });

  test("gives each stroke pattern a distinguishable background", () => {
    const patterns: readonly DisplayStrokePattern[] = [
      "solid",
      "dashed",
      "dotted",
      "double",
      "wavy",
    ];

    const backgrounds = patterns.map((pattern) => {
      const element = renderPrimitives([
        { ...PRIMITIVE_SAMPLES.line, stroke: { color: BLUE, thicknessPx: 2, pattern } },
      ]).children.at(0);
      return element?.style.background ?? "";
    });

    expect(new Set(backgrounds).size).toBe(patterns.length);
    expect(backgrounds.every((background) => background.length > 0)).toBe(true);
  });

  test("reaches across the path only as far as the pattern needs", () => {
    const crossExtent = (pattern: DisplayStrokePattern) => {
      const element = renderPrimitives([
        { ...PRIMITIVE_SAMPLES.line, stroke: { color: BLUE, thicknessPx: 2, pattern } },
      ]).children.at(0);
      return [element?.style.height, element?.style.top];
    };

    // Centred on the path: a 2px dash sits 1px above y=10, a 6px double 3px.
    expect(crossExtent("dashed")).toEqual(["2px", "9px"]);
    expect(crossExtent("double")).toEqual(["6px", "7px"]);
    expect(crossExtent("wavy")).toEqual(["6px", "7px"]);
  });

  test("draws dashes on the shared period", () => {
    const element = renderPrimitives([
      { ...PRIMITIVE_SAMPLES.line, stroke: { color: BLUE, thicknessPx: 2, pattern: "dashed" } },
    ]).children.at(0);

    // STROKE_DASH_FACTORS.dashed is 3 on, 2 off, at 2px thick.
    expect(element?.style.background).toBe(
      "repeating-linear-gradient(to right, rgb(0, 0, 255) 0px, rgb(0, 0, 255) 6px, transparent 6px, transparent 10px)",
    );
  });

  test("rotates a diagonal line about its first endpoint", () => {
    const element = renderPrimitives([
      { ...PRIMITIVE_SAMPLES.line, x1Px: 0, y1Px: 0, x2Px: 3, y2Px: 4 },
    ]).children.at(0);

    expect(element?.style.width).toBe("5px");
    expect(element?.style.transformOrigin).toBe("0 50%");
    expect(element?.style.transform).toBe("rotate(53.13010235415598deg)");
  });

  test("clips a group and offsets its children to the clip origin", () => {
    const clip = renderPrimitives([
      {
        kind: "clipGroup",
        rect: { xPx: 10, yPx: 20, widthPx: 100, heightPx: 50 },
        children: [
          {
            kind: "rect",
            rect: { xPx: 30, yPx: 60, widthPx: 5, heightPx: 5 },
            fill: BLUE,
          },
        ],
      },
    ]).children.at(0);

    expect(clip?.style.overflow).toBe("hidden");
    expect(clip?.style.left).toBe("10px");
    expect(clip?.style.top).toBe("20px");
    expect(clip?.style.width).toBe("100px");

    const child = clip?.children.at(0);
    expect(child?.style.left).toBe("20px");
    expect(child?.style.top).toBe("40px");
  });

  test("keeps child coordinates through rotate and opacity groups", () => {
    const rotate = renderPrimitives([
      {
        kind: "rotateGroup",
        degrees: 90,
        originXPx: 40,
        originYPx: 50,
        children: [PRIMITIVE_SAMPLES.rect],
      },
    ]).children.at(0);

    expect(rotate?.style.transform).toBe("rotate(90deg)");
    expect(rotate?.style.transformOrigin).toBe("40px 50px");
    expect(rotate?.children.at(0)?.style.left).toBe("10px");

    const opacity = renderPrimitives([
      { kind: "opacityGroup", opacity: 0.25, children: [PRIMITIVE_SAMPLES.rect] },
    ]).children.at(0);

    expect(opacity?.style.opacity).toBe("0.25");
    expect(opacity?.children.at(0)?.style.top).toBe("20px");
  });

  test("serializes an opaque colour as rgb and a translucent one as rgba", () => {
    const opaque = renderPrimitives([
      { kind: "rect", rect: PRIMITIVE_SAMPLES.rect.rect, fill: { r: 1, g: 2, b: 3, a: 1 } },
    ]).children.at(0);
    const translucent = renderPrimitives([
      { kind: "rect", rect: PRIMITIVE_SAMPLES.rect.rect, fill: { r: 1, g: 2, b: 3, a: 0.5 } },
    ]).children.at(0);

    expect(opaque?.style.backgroundColor).toBe("rgb(1, 2, 3)");
    expect(translucent?.style.backgroundColor).toBe("rgba(1, 2, 3, 0.5)");
  });

  test("centres a solid rect border on the rect path", () => {
    const element = renderPrimitives([
      {
        kind: "rect",
        rect: { xPx: 10, yPx: 20, widthPx: 30, heightPx: 40 },
        stroke: { color: BLUE, thicknessPx: 2, pattern: "solid" },
      },
    ]).children.at(0);

    expect(element?.style.left).toBe("9px");
    expect(element?.style.width).toBe("32px");
    expect(element?.style.boxSizing).toBe("border-box");
    expect(element?.style.border).toBe("2px solid rgb(0, 0, 255)");
  });

  test("strokes a patterned rect edge by edge, sharing the line implementation", () => {
    for (const pattern of ["dashed", "dotted", "double", "wavy"] as const) {
      const children = renderPrimitives([
        {
          kind: "rect",
          rect: { xPx: 0, yPx: 0, widthPx: 10, heightPx: 10 },
          stroke: { color: BLUE, thicknessPx: 1, pattern },
        },
      ]).children;

      expect(children, pattern).toHaveLength(5);
      expect(children.at(0)?.style.border, pattern).toBeUndefined();
      const line = renderPrimitives([
        {
          kind: "line",
          x1Px: 0,
          y1Px: 0,
          x2Px: 10,
          y2Px: 0,
          stroke: { color: BLUE, thicknessPx: 1, pattern },
        },
      ]).children.at(0);
      expect(children.at(1)?.style.background, pattern).toBe(line?.style.background ?? "");
    }
  });

  test("crops an image by oversizing it inside a clipping div", () => {
    const clip = renderPrimitives([
      {
        kind: "image",
        image: 0,
        rect: { xPx: 0, yPx: 0, widthPx: 50, heightPx: 50 },
        crop: { l: 0.25, t: 0, r: 0.25, b: 0.5 },
        opacity: 0.5,
      },
    ]).children.at(0);

    expect(clip?.style.overflow).toBe("hidden");

    const image = clip?.children.at(0);
    expect(image?.tagName).toBe("img");
    expect(image?.style.width).toBe("100px");
    expect(image?.style.height).toBe("100px");
    expect(image?.style.left).toBe("-25px");
    expect(image?.style.top).toBe("0px");
    expect(image?.style.opacity).toBe("0.5");
    expect(image?.src.startsWith("blob:") || image?.src.startsWith("data:image/png")).toBe(true);
  });

  test("registers an embedded face once and names it ahead of the family", () => {
    const pages = renderDisplayListToDom(
      listOf([emptyPage([PRIMITIVE_SAMPLES.glyphRun]), emptyPage([])], [EMBEDDED_FONT]),
      { doc: stubDocument() },
    ).map(asStub);

    const rules = pages.map((page) => page.children.at(0));
    expect(rules.map((rule) => rule?.tagName)).toEqual(["style", "style"]);
    expect(rules.at(0)?.textContent).toContain(`font-family: "embedded-1"`);
    expect(rules.at(0)?.textContent).toContain("font-weight: 400");
    // One blob for the whole render, so one revoke releases the bytes.
    expect(rules.at(1)?.textContent).toBe(rules.at(0)?.textContent ?? "");

    const span = pages.at(0)?.children.at(1);
    expect(span?.style.fontFamily).toBe(`"embedded-1", "Folio Sans", sans-serif`);
  });

  test("omits the style element when no face is embedded", () => {
    expect(renderPrimitives([]).children).toHaveLength(0);
  });

  test("paints links above the marks, addressing pages by list index", () => {
    const page = asStub(
      renderDisplayPageToDom(
        {
          ...emptyPage([PRIMITIVE_SAMPLES.rect]),
          links: [
            {
              rect: { xPx: 1, yPx: 2, widthPx: 3, heightPx: 4 },
              target: { kind: "external", href: "https://example.org/" },
              tooltip: "Example",
            },
            {
              rect: { xPx: 5, yPx: 6, widthPx: 7, heightPx: 8 },
              target: { kind: "page", pageIndex: 2, yPx: 100 },
            },
          ],
        },
        { doc: stubDocument(), fonts: [FONT], images: [IMAGE], pageIndex: 0 },
      ),
    );

    const [external, internal] = page.children.slice(1);
    expect(external?.tagName).toBe("a");
    expect(external?.href).toBe("https://example.org/");
    expect(external?.title).toBe("Example");
    expect(external?.style.left).toBe("1px");
    expect(internal?.href).toBe("#page-2");
  });

  test("panics on a font ref the display list cannot resolve", () => {
    expect(() =>
      renderDisplayPageToDom(emptyPage([PRIMITIVE_SAMPLES.glyphRun]), {
        doc: stubDocument(),
        fonts: [],
        images: [],
        pageIndex: 0,
      }),
    ).toThrow(/font ref 0 is out of range/);
  });

  test("paints the page background before any primitive", () => {
    const page = asStub(
      renderDisplayPageToDom(emptyPage([]), {
        doc: stubDocument(),
        fonts: [],
        images: [],
        pageIndex: 0,
        pageBackground: { r: 255, g: 255, b: 255, a: 1 },
      }),
    );

    expect(page.style.backgroundColor).toBe("rgb(255, 255, 255)");
  });
});
