// Toggling a bullet / numbered list while editing a header must show a marker
// on the painted header and survive the save round-trip. Lists in folio are
// modelled as paragraph attributes (numPr / listMarker), not list nodes, so the
// path is: toggle command sets numPr on the HF paragraph -> the PM-doc HF
// converter resolves the marker glyph -> renderPage paints it via
// renderParagraphFragment -> proseDocToBlocks + the HeaderFooter converter keep
// it across save/reload. This locks every link so an HF list regression fails
// here instead of silently painting an empty header.

import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { clearTextWidthCache } from "../layout-engine/measure/cache";
import { resetCanvasContext } from "../layout-engine/measure/measureContainer";
import { measureBlocks as realMeasureBlocks } from "../layout-engine/measure/measureBlocks";
import type { FlowBlock, Measure, Page } from "../layout-engine/types";
import {
  convertHeaderFooterPmDocToContent,
  convertHeaderFooterToContent,
} from "../layout-bridge/convert/headerFooterLayout";
import type { HeaderFooterMetrics } from "../layout-bridge/convert/headerFooterLayout";
import { toggleBulletList, toggleNumberedList } from "../prosemirror/commands/paragraph";
import { headerFooterToProseDoc } from "../prosemirror/conversion/toProseDoc";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { schema, singletonManager } from "../prosemirror/schema";
import type { HeaderFooter } from "../types/document";
import { renderPage } from "../layout-painter/renderPage";

function createFakeStyle(): Record<string, string> {
  const store: Record<string, string> = {};
  return new Proxy(store, {
    get(target, prop: string) {
      if (prop === "setProperty") {
        return (key: string, value: string) => {
          target[key] = value;
        };
      }
      if (prop === "getPropertyValue") {
        return (key: string) => target[key] ?? "";
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
  innerHTML = "";
  dir = "";
  textContent = "";
  style = createFakeStyle();
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
  prepend(...children: FakeElement[]): void {
    this.children.unshift(...children);
  }
  appendChild(child: FakeElement): FakeElement {
    this.children.push(child);
    return child;
  }
  querySelector(): FakeElement | null {
    return null;
  }
  querySelectorAll(): FakeElement[] {
    return [];
  }
  getContext(): { font: string; measureText: (text: string) => { width: number } } | null {
    if (this.tagName !== "canvas") {
      return null;
    }
    return { font: "", measureText: (t: string) => ({ width: t.length * 7 }) };
  }
}

const fakeDocument = {
  createElement: (tag: string) => new FakeElement(tag),
} as unknown as Document;

const metrics: HeaderFooterMetrics = {
  section: "header",
  pageSize: { w: 600, h: 800 },
  margins: { top: 100, right: 72, bottom: 100, left: 72, header: 48, footer: 48 },
};

const page = {
  number: 1,
  size: { w: 600, h: 800 },
  margins: { top: 100, right: 72, bottom: 100, left: 72, header: 48, footer: 48 },
  fragments: [],
} as unknown as Page;

const hfMeasureBlocks = (blocks: FlowBlock[], width: number): Measure[] =>
  realMeasureBlocks(blocks, width);

function collectMarkers(root: FakeElement): FakeElement[] {
  const out: FakeElement[] = [];
  const walk = (el: FakeElement) => {
    if (el.className.includes("layout-list-marker")) {
      out.push(el);
    }
    for (const child of el.children) {
      walk(child);
    }
  };
  walk(root);
  return out;
}

// Build an HF PM state the way HiddenHeaderFooterPMs.buildInitialState does
// (headerFooterToProseDoc + the singleton starter-kit schema/plugins), place a
// caret in the first paragraph, and run the toggle command the toolbar
// dispatches. Returns the mutated doc.
function toggleListInHeader(command: typeof toggleBulletList, headerText: string) {
  const hf: HeaderFooter = {
    type: "header",
    hdrFtrType: "default",
    content: [
      {
        type: "paragraph",
        content: headerText ? [{ type: "run", content: [{ type: "text", text: headerText }] }] : [],
      },
    ],
  };
  const pmDoc = headerFooterToProseDoc(hf.content);
  let state = EditorState.create({ doc: pmDoc, schema, plugins: singletonManager.getPlugins() });
  state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1)));
  let next: EditorState | null = null;
  const handled = command(state, (tr) => {
    next = state.apply(tr);
  });
  if (!handled || !next) {
    throw new Error("toggle command did not dispatch");
  }
  return (next as EditorState).doc;
}

function withFakeDocument<T>(run: () => T): T {
  const original = globalThis.document;
  Object.defineProperty(globalThis, "document", { value: fakeDocument, configurable: true });
  clearTextWidthCache();
  resetCanvasContext();
  try {
    return run();
  } finally {
    clearTextWidthCache();
    resetCanvasContext();
    Object.defineProperty(globalThis, "document", { value: original, configurable: true });
  }
}

describe("header/footer list toggling", () => {
  test("toggleBulletList sets numPr on the header paragraph", () => {
    const doc = toggleListInHeader(toggleBulletList, "Header line");
    const para = doc.child(0);
    expect(para.type.name).toBe("paragraph");
    expect(para.attrs["numPr"]).toEqual({ numId: 1, ilvl: 0 });
    expect(para.attrs["listIsBullet"]).toBe(true);
  });

  test("bullet toggle paints a marker on the header (with and without text)", () => {
    for (const headerText of ["Header line", ""]) {
      withFakeDocument(() => {
        const doc = toggleListInHeader(toggleBulletList, headerText);
        const headerContent = convertHeaderFooterPmDocToContent(doc, 456, metrics, {
          measureBlocks: hfMeasureBlocks,
        });
        const block = headerContent?.blocks[0];
        expect(block?.kind === "paragraph" ? block.attrs?.listMarker : null).toBe("•");

        const pageEl = renderPage(
          page,
          { pageNumber: 1, totalPages: 1, section: "body" },
          { document: fakeDocument, headerContent },
        ) as unknown as FakeElement;
        const markers = collectMarkers(pageEl);
        expect(markers.length).toBe(1);
        expect(markers[0]?.textContent).toBe("•");
      });
    }
  });

  test("numbered toggle paints a decimal marker on the header", () => {
    withFakeDocument(() => {
      const doc = toggleListInHeader(toggleNumberedList, "Header line");
      const headerContent = convertHeaderFooterPmDocToContent(doc, 456, metrics, {
        measureBlocks: hfMeasureBlocks,
      });
      const block = headerContent?.blocks[0];
      expect(block?.kind === "paragraph" ? block.attrs?.listMarker : null).toBe("1.");
      expect(block?.kind === "paragraph" ? block.attrs?.listIsBullet : null).toBe(false);

      const pageEl = renderPage(
        page,
        { pageNumber: 1, totalPages: 1, section: "body" },
        { document: fakeDocument, headerContent },
      ) as unknown as FakeElement;
      const markers = collectMarkers(pageEl);
      expect(markers.length).toBe(1);
      expect(markers[0]?.textContent).toBe("1.");
    });
  });

  test("a toggled list survives the save round-trip into the header", () => {
    withFakeDocument(() => {
      const doc = toggleListInHeader(toggleBulletList, "Header line");
      // Save flushes the live PM doc back to HeaderFooter.content via proseDocToBlocks.
      const savedBlocks = proseDocToBlocks(doc);
      const savedParagraph = savedBlocks.find((b) => b.type === "paragraph");
      expect(
        savedParagraph?.type === "paragraph" ? savedParagraph.formatting?.numPr : null,
      ).toEqual({ numId: 1, ilvl: 0 });

      // Reload paints from HeaderFooter.content; the bullet marker is recomputed.
      const savedHf: HeaderFooter = { type: "header", hdrFtrType: "default", content: savedBlocks };
      const reloaded = convertHeaderFooterToContent(savedHf, 456, metrics, {
        measureBlocks: hfMeasureBlocks,
      });
      const reloadedParagraph = reloaded?.blocks.find((b) => b.kind === "paragraph");
      expect(
        reloadedParagraph?.kind === "paragraph" ? reloadedParagraph.attrs?.listMarker : null,
      ).toBe("•");
    });
  });
});
