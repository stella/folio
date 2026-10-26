import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

import { panic } from "better-result";
import { afterAll, afterEach, beforeEach, expect, test } from "bun:test";
import type { EditorView } from "prosemirror-view";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { sizeProportionalCharWidth } from "@stll/folio-core/layout-engine/measure/__tests__/fakeTextMeasure";
import { clearAllCaches } from "@stll/folio-core/layout-engine/measure/cache";
import { resetCanvasContext } from "@stll/folio-core/layout-engine/measure/measureContainer";
import type { Paragraph } from "@stll/folio-core/types/document";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import type { FolioOutlineRailProps, OutlineItem } from "../ui/folio-ui";
import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const SECTION_COUNT = 4;
const CLAUSES_PER_SECTION = 10;
/** Layout with measured glyphs is slow under happy-dom. */
const EDITOR_TEST_TIMEOUT_MS = 60_000;
/** Past the editor's former 100/300/400 ms outline timers. */
const SETTLE_MS = 700;

const CLAUSE =
  "The Supplier shall deliver the Services with reasonable skill and care, keep complete records of the work performed, and give the Customer prompt written notice of any event that may delay performance.";

const paragraph = (text: string, styleId?: string): Paragraph => ({
  type: "paragraph",
  formatting: styleId ? { styleId } : {},
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const createLongDocument = () => {
  const document = createEmptyDocument();
  const content: Paragraph[] = [];
  for (let section = 1; section <= SECTION_COUNT; section++) {
    content.push(paragraph(`Article ${String(section)}`, "Heading1"));
    for (let clause = 0; clause < CLAUSES_PER_SECTION; clause++) {
      content.push(paragraph(clause === 0 ? "Short clause." : CLAUSE));
    }
  }
  document.package.document.content = content;
  return document;
};

/** Every item list the outline rail was rendered with, in order. */
const renders: OutlineItem[][] = [];
const ProbeOutlineRail = ({ items }: FolioOutlineRailProps) => {
  renders.push(items);
  return null;
};
const PROBE_COMPONENTS = { OutlineRail: ProbeOutlineRail };

const previousClientWidth = Object.getOwnPropertyDescriptor(
  HTMLElement.prototype,
  "clientWidth",
);
const originalGetContext =
  Object.getOwnPropertyDescriptor(HTMLCanvasElement.prototype, "getContext") ??
  panic("Expected the canvas context method");
// Deterministic glyph widths, so the document paginates without a browser.
const canvasContext = {
  font: "",
  fontKerning: "auto",
  measureText(text: string) {
    let width = 0;
    for (const char of text)
      width += sizeProportionalCharWidth(char, this.font);
    return { width, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 };
  },
};
beforeEach(() => {
  renders.length = 0;
  // Wide enough for the outline column.
  Object.defineProperty(HTMLElement.prototype, "clientWidth", {
    configurable: true,
    get(this: HTMLElement) {
      if (this.classList.contains("folio-panels-row")) return 1440;
      return previousClientWidth?.get?.call(this) ?? 0;
    },
  });
  Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
    configurable: true,
    value: () => canvasContext,
  });
  resetCanvasContext();
  clearAllCaches();
});
afterEach(() => {
  Object.defineProperty(
    HTMLCanvasElement.prototype,
    "getContext",
    originalGetContext,
  );
  resetCanvasContext();
  clearAllCaches();
  if (previousClientWidth) {
    Object.defineProperty(
      HTMLElement.prototype,
      "clientWidth",
      previousClientWidth,
    );
  } else {
    Reflect.deleteProperty(HTMLElement.prototype, "clientWidth");
  }
});

const settle = () =>
  act(async () => {
    await new Promise((resolve) => setTimeout(resolve, SETTLE_MS));
  });

const mountEditor = async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  await act(async () =>
    root.render(
      <IntlProvider
        locale="en"
        timeZone="UTC"
        messages={getFolioMessages("en")}
      >
        <DocxEditor
          ref={editor}
          document={createLongDocument()}
          components={PROBE_COMPONENTS}
          showToolbar={false}
        />
      </IntlProvider>,
    ),
  );
  await act(async () => editor.current?.ensureEditorView({ focus: false }));
  await settle();
  const view = (): EditorView =>
    editor.current?.getEditor().getView() ??
    panic("The editor did not create its body view");
  const latest = () =>
    renders.at(-1) ?? panic("The outline rail never rendered");
  return {
    view,
    latest,
    unmount: async () => {
      await act(async () => root.unmount());
      container.remove();
    },
  };
};

/** Document position of the paragraph whose text is `text`. */
const paragraphPos = (view: EditorView, text: string): number => {
  let found: number | null = null;
  view.state.doc.descendants((node, pos) => {
    if (
      found === null &&
      node.type.name === "paragraph" &&
      node.textContent === text
    ) {
      found = pos;
    }
    return found === null;
  });
  return found ?? panic(`No paragraph reads "${text}"`);
};

const pageNumbers = (items: readonly OutlineItem[]) =>
  items.map(({ meta }) => meta);

/** What the rail was handed after an edit: diagnostic, logged before asserting. */
const reportRenders = (edit: string, since: number) => {
  const after = renders.slice(since);
  console.info(
    "FOLIO_OUTLINE_EDIT_MEASUREMENT",
    JSON.stringify({
      edit,
      railRenders: after.length,
      newItemLists: after.filter(
        (items, index) => items !== renders[since + index - 1],
      ).length,
      entriesWithoutPage: after.reduce(
        (count, items) =>
          count + items.filter(({ meta }) => meta === undefined).length,
        0,
      ),
    }),
  );
  return after;
};

test(
  "an edit that moves no heading leaves the outline entries as they were",
  async () => {
    const { view, latest, unmount } = await mountEditor();
    try {
      const before = latest();
      expect(before).toHaveLength(SECTION_COUNT);
      expect(new Set(pageNumbers(before)).size).toBeGreaterThan(1);
      const rendersBefore = renders.length;

      await act(async () => {
        const from = paragraphPos(view(), "Short clause.") + 1;
        view().dispatch(
          view().state.tr.addMark(
            from,
            from + "Short".length,
            view().state.schema.marks["bold"]?.create() ??
              panic("The schema has no bold mark"),
          ),
        );
      });
      await settle();

      expect(
        reportRenders("bold", rendersBefore).filter(
          (items) => items !== before,
        ),
      ).toEqual([]);
    } finally {
      await unmount();
    }
  },
  EDITOR_TEST_TIMEOUT_MS,
);

test(
  "typing inside a heading keeps every entry's id and page number",
  async () => {
    const { view, latest, unmount } = await mountEditor();
    try {
      const before = latest();
      const rendersBefore = renders.length;

      await act(async () => {
        view().dispatch(
          view().state.tr.insertText(
            "Draft ",
            paragraphPos(view(), "Article 1") + 1,
          ),
        );
      });
      await settle();

      const rendered = reportRenders("type in a heading", rendersBefore);
      const after = latest();
      expect(after.map(({ id }) => id)).toEqual(before.map(({ id }) => id));
      expect(pageNumbers(after)).toEqual(pageNumbers(before));
      expect(after.at(0)?.label).toBe("Draft Article 1");
      for (const items of rendered) {
        expect(items.every(({ meta }) => meta !== undefined)).toBe(true);
      }
    } finally {
      await unmount();
    }
  },
  EDITOR_TEST_TIMEOUT_MS,
);

test(
  "a heading that moves to a later page changes only its own page number",
  async () => {
    const { view, latest, unmount } = await mountEditor();
    try {
      const before = latest();
      const rendersBefore = renders.length;
      const lastHeading = `Article ${String(SECTION_COUNT)}`;

      await act(async () => {
        const at = paragraphPos(view(), lastHeading);
        const filler = Array.from(
          { length: CLAUSES_PER_SECTION },
          () =>
            view().state.schema.nodes["paragraph"]?.create(
              null,
              view().state.schema.text(CLAUSE),
            ) ?? panic("The schema has no paragraph node"),
        );
        view().dispatch(view().state.tr.insert(at, filler));
      });
      await settle();

      const rendered = reportRenders("move the last heading", rendersBefore);
      const after = latest();
      expect(after.map(({ id }) => id)).toEqual(before.map(({ id }) => id));
      expect(pageNumbers(after).slice(0, -1)).toEqual(
        pageNumbers(before).slice(0, -1),
      );
      expect(Number(after.at(-1)?.meta)).toBeGreaterThan(
        Number(before.at(-1)?.meta),
      );
      for (const items of rendered) {
        expect(items.every(({ meta }) => meta !== undefined)).toBe(true);
      }
    } finally {
      await unmount();
    }
  },
  EDITOR_TEST_TIMEOUT_MS,
);
