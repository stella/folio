/**
 * Text-box module contract test. Verifies kind discriminator and that
 * style attributes (fill, border, padding) flow through the module.
 */

import { describe, expect, test } from "bun:test";

import type { TextBoxBlock, TextBoxFragment, TextBoxMeasure } from "../../../layout-engine/types";
import { cssLinearGradientAngle, TEXTBOX_CLASS_NAMES } from "../../renderTextBox";
import type { RenderContext } from "../../renderUtils";
import { textBoxModule } from "./textBox";

function createFakeStyle(): Record<string, string> {
  const store: Record<string, string> = {};
  return new Proxy(store, {
    get(target, prop: string) {
      if (prop === "setProperty") {
        return (key: string, value: string) => {
          target[key] = value;
        };
      }
      return target[prop];
    },
    set(target, prop: string, value: string) {
      target[prop] = value;
      return true;
    },
  }) as unknown as Record<string, string>;
}

class FakeElement {
  className = "";
  dataset: Record<string, string> = {};
  style: Record<string, string> = createFakeStyle();
  children: FakeElement[] = [];
  classList = {
    add: (...tokens: string[]) => {
      this.className = [this.className, ...tokens].filter(Boolean).join(" ");
    },
  };
  readonly tagName: string;

  constructor(tagName: string) {
    this.tagName = tagName;
  }

  append(...children: FakeElement[]): void {
    this.children.push(...children);
  }
}

const fakeDocument = {
  createElement(tag: string): FakeElement {
    return new FakeElement(tag);
  },
} as unknown as Document;

const ctx: RenderContext = {
  pageNumber: 1,
  totalPages: 1,
  section: "body",
};

describe("textBoxModule", () => {
  test("identifies itself as the textBox kind", () => {
    expect(textBoxModule.kind).toBe("textBox");
  });

  test.each([
    ["sysDash", "dashed"],
    ["lgDashDot", "dashed"],
    ["dot", "dotted"],
    ["dashed", "dashed"],
  ] as const)("paints a %s outline as a CSS %s border", (outlineStyle, cssStyle) => {
    // `outlineStyle` is a DrawingML dash, not a CSS keyword: interpolating it
    // into the shorthand made the whole declaration invalid and the outline
    // vanished.
    const el = textBoxModule.render({
      fragment: { kind: "textBox", blockId: "tb-dash", x: 0, y: 0, width: 200, height: 100 },
      block: {
        kind: "textBox",
        id: "tb-dash",
        width: 200,
        height: 100,
        outlineWidth: 2,
        outlineColor: "var(--test-outline)",
        outlineStyle,
        content: [],
      },
      measure: { kind: "textBox", width: 200, height: 100, innerMeasures: [] },
      context: ctx,
      doc: fakeDocument,
    }) as unknown as FakeElement;

    expect(el.style["border"]).toBe(`2px ${cssStyle} var(--test-outline)`);
  });

  test("paints no border for the explicit no-outline sentinel", () => {
    const el = textBoxModule.render({
      fragment: { kind: "textBox", blockId: "tb-none", x: 0, y: 0, width: 200, height: 100 },
      block: {
        kind: "textBox",
        id: "tb-none",
        width: 200,
        height: 100,
        outlineWidth: 2,
        outlineColor: "var(--test-outline)",
        outlineStyle: "none",
        content: [],
      },
      measure: { kind: "textBox", width: 200, height: 100, innerMeasures: [] },
      context: ctx,
      doc: fakeDocument,
    }) as unknown as FakeElement;

    expect(el.style["border"]).toBeUndefined();
  });

  test("renders a text box fragment with fill, border, and padding", () => {
    const fragment: TextBoxFragment = {
      kind: "textBox",
      blockId: "tb1",
      x: 0,
      y: 0,
      width: 200,
      height: 100,
    };
    const block: TextBoxBlock = {
      kind: "textBox",
      id: "tb1",
      width: 200,
      height: 100,
      fillColor: "var(--test-fill)",
      outlineWidth: 2,
      outlineColor: "var(--test-outline)",
      outlineStyle: "solid",
      margins: { top: 4, right: 6, bottom: 4, left: 6 },
      content: [],
    };
    const measure: TextBoxMeasure = {
      kind: "textBox",
      width: 200,
      height: 100,
      innerMeasures: [],
    };

    const el = textBoxModule.render({
      fragment,
      block,
      measure,
      context: ctx,
      doc: fakeDocument,
    }) as unknown as FakeElement;

    expect(el.className).toBe(TEXTBOX_CLASS_NAMES.textBox);
    expect(el.style["backgroundColor"]).toBe("var(--test-fill)");
    expect(el.style["border"]).toBe("2px solid var(--test-outline)");
    expect(el.style["padding"]).toBe("4px 6px 4px 6px");
    expect(el.dataset["blockId"]).toBe("tb1");
  });

  test("paints a scaled linear gradient corner to corner of a tall box", () => {
    // 315° in the unit square runs from the bottom-left corner to the top-right
    // one. Stretched to 100x400 the isolines stay parallel to the stretched
    // diagonal, so the gradient itself turns towards the horizontal.
    const el = textBoxModule.render({
      fragment: { kind: "textBox", blockId: "tb-grad", x: 0, y: 0, width: 100, height: 400 },
      block: {
        kind: "textBox",
        id: "tb-grad",
        width: 100,
        height: 400,
        fillGradient: {
          angle: 315,
          scaled: true,
          stops: [
            { offset: 0, color: "#112233" },
            { offset: 0.5, color: "#445566" },
            { offset: 1, color: "#778899" },
          ],
        },
        content: [],
      },
      measure: { kind: "textBox", width: 100, height: 400, innerMeasures: [] },
      context: ctx,
      doc: fakeDocument,
    }) as unknown as FakeElement;

    const expectedAngle = 90 + (Math.atan2(-100, 400) * 180) / Math.PI;
    expect(el.style["backgroundImage"]).toBe(
      `linear-gradient(${Math.round(expectedAngle * 1000) / 1000}deg, #112233 0%, #445566 50%, #778899 100%)`,
    );
    expect(el.style["backgroundColor"]).toBeUndefined();
  });

  test.each([
    [0, false, 100, 400, 90],
    [90, false, 100, 400, 180],
    [315, false, 100, 400, 45],
    [45, true, 200, 200, 135],
    [0, true, 100, 400, 90],
  ] as const)(
    "turns a %s° gradient (scaled: %s) over %sx%s into a %s° CSS angle",
    (angle, scaled, width, height, cssAngle) => {
      expect(cssLinearGradientAngle({ angle, scaled }, width, height)).toBeCloseTo(cssAngle, 6);
    },
  );

  test("a solid fill color wins over a gradient", () => {
    const el = textBoxModule.render({
      fragment: { kind: "textBox", blockId: "tb-both", x: 0, y: 0, width: 100, height: 100 },
      block: {
        kind: "textBox",
        id: "tb-both",
        width: 100,
        height: 100,
        fillColor: "#ABCDEF",
        fillGradient: {
          angle: 0,
          scaled: false,
          stops: [
            { offset: 0, color: "#000000" },
            { offset: 1, color: "#FFFFFF" },
          ],
        },
        content: [],
      },
      measure: { kind: "textBox", width: 100, height: 100, innerMeasures: [] },
      context: ctx,
      doc: fakeDocument,
    }) as unknown as FakeElement;

    expect(el.style["backgroundColor"]).toBe("#ABCDEF");
    expect(el.style["backgroundImage"]).toBeUndefined();
  });

  test.each([
    ["middle", "center"],
    ["bottom", "flex-end"],
  ] as const)("renders %s-aligned text box content", (verticalAlign, justifyContent) => {
    const fragment: TextBoxFragment = {
      kind: "textBox",
      blockId: `tb-${verticalAlign}`,
      x: 0,
      y: 0,
      width: 200,
      height: 100,
    };
    const block: TextBoxBlock = {
      kind: "textBox",
      id: `tb-${verticalAlign}`,
      width: 200,
      height: 100,
      verticalAlign,
      content: [],
    };
    const measure: TextBoxMeasure = {
      kind: "textBox",
      width: 200,
      height: 100,
      innerMeasures: [],
    };

    const el = textBoxModule.render({
      fragment,
      block,
      measure,
      context: ctx,
      doc: fakeDocument,
    }) as unknown as FakeElement;

    expect(el.style["display"]).toBe("flex");
    expect(el.style["flexDirection"]).toBe("column");
    expect(el.style["justifyContent"]).toBe(justifyContent);
  });

  test.each(["distributed", "justified"] as const)(
    "does not approximate %s line alignment with block flex spacing",
    (verticalAlign) => {
      const fragment: TextBoxFragment = {
        kind: "textBox",
        blockId: `tb-${verticalAlign}`,
        x: 0,
        y: 0,
        width: 200,
        height: 100,
      };
      const block: TextBoxBlock = {
        kind: "textBox",
        id: `tb-${verticalAlign}`,
        width: 200,
        height: 100,
        verticalAlign,
        content: [],
      };
      const measure: TextBoxMeasure = {
        kind: "textBox",
        width: 200,
        height: 100,
        innerMeasures: [],
      };

      const el = textBoxModule.render({
        fragment,
        block,
        measure,
        context: ctx,
        doc: fakeDocument,
      }) as unknown as FakeElement;

      expect(el.style["display"]).toBeUndefined();
      expect(el.style["justifyContent"]).toBeUndefined();
    },
  );

  test("renders nested tables in text box content", () => {
    const fragment: TextBoxFragment = {
      kind: "textBox",
      blockId: "tb-table",
      x: 0,
      y: 0,
      width: 200,
      height: 100,
    };
    const block: TextBoxBlock = {
      kind: "textBox",
      id: "tb-table",
      width: 200,
      height: 100,
      margins: { top: 0, right: 0, bottom: 0, left: 0 },
      content: [
        {
          kind: "table",
          id: "nested-table",
          rows: [],
        },
      ],
    };
    const measure: TextBoxMeasure = {
      kind: "textBox",
      width: 200,
      height: 100,
      innerMeasures: [
        {
          kind: "table",
          rows: [],
          columnWidths: [],
          totalWidth: 160,
          totalHeight: 20,
        },
      ],
    };

    const el = textBoxModule.render({
      fragment,
      block,
      measure,
      context: ctx,
      doc: fakeDocument,
    }) as unknown as FakeElement;

    expect(el.children).toHaveLength(1);
    expect(el.children.at(0)?.className).toContain("layout-table");
    expect(el.children.at(0)?.dataset["blockId"]).toBe("nested-table");
  });
});
