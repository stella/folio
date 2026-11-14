/**
 * The Vue editor reads the model it holds and never writes it: the hidden
 * editor state, the layout pipeline, the initial font wait, the header/footer
 * and note editors' style readers and the change write-back all take the model
 * as it is. This file loads every document shape deep-frozen and drives those
 * paths, from the precomputed first paint through an edit and a save. None
 * may throw on the frozen model, and each must give what it gives for the
 * same document unfrozen.
 */

import { GlobalRegistrator } from "@happy-dom/global-registrator";

// PM keymaps capture the host platform at import time. Keep the DOM fixture on
// that same platform when another test has already imported the editor schema.
const hostPlatform = typeof navigator === "undefined" ? "" : navigator.platform;
GlobalRegistrator.register();
Object.defineProperty(navigator, "platform", {
  configurable: true,
  value: hostPlatform,
});

import { panic } from "better-result";
import {
  afterAll,
  afterEach,
  beforeEach,
  expect,
  setDefaultTimeout,
  test,
} from "bun:test";

const { createApp, defineComponent, h, shallowRef } = await import("vue");

import { deepFreeze } from "@stll/folio-core/__tests__/deepFreeze";
import { DOCUMENT_SHAPES } from "@stll/folio-core/__tests__/documentShapes";
import { DOCUMENT_CHANGE_NOTIFY_DELAY_MS } from "@stll/folio-core/controller/documentChangeNotification";
import { prepareCanonicalDocxInput } from "@stll/folio-core/docx/canonicalSessionInput";
import { parseDocx } from "@stll/folio-core/docx/parser";
import type {} from "@stll/folio-core/layout-engine/layoutInstrumentation";
import { sizeProportionalCharWidth } from "@stll/folio-core/layout-engine/measure/__tests__/fakeTextMeasure";
import { clearAllCaches } from "@stll/folio-core/layout-engine/measure/cache";
import { resetCanvasContext } from "@stll/folio-core/layout-engine/measure/measureContainer";
import type { Document } from "@stll/folio-core/types/document";

const { useDocxEditor } = await import("./useDocxEditor");

setDefaultTimeout(60_000);

afterAll(() => GlobalRegistrator.unregister());

const originalGetContext =
  Object.getOwnPropertyDescriptor(HTMLCanvasElement.prototype, "getContext") ??
  panic("Expected the canvas context method");
// Deterministic glyph widths, so every shape lays out without a browser.
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
});

const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));
/** Past the initial font wait and the first layout pass. */
const paint = () => wait(30);
/** Past the debounced change notification. */
const settle = () => wait(DOCUMENT_CHANGE_NOTIFY_DELAY_MS + 100);

const parseShape = async (bytes: Uint8Array): Promise<Document> =>
  parseDocx((await prepareCanonicalDocxInput(bytes.slice().buffer)).unwrap(), {
    preloadFonts: false,
    detectVariables: false,
  });

const typeInto = (dom: HTMLElement, data: string) =>
  dom.dispatchEvent(
    new InputEvent("beforeinput", {
      bubbles: true,
      cancelable: true,
      inputType: "insertText",
      data,
    }),
  );

/**
 * Loads `document` into a default-session Vue editor before its host mounts
 * (the precomputed first paint), mounts the host (the view's initial layout
 * and font wait), then reads, lays out, types and saves, and returns what
 * each path produced.
 */
const driveEditor = async (document: Document) => {
  const container = window.document.createElement("div");
  const hidden = window.document.createElement("div");
  const pages = window.document.createElement("div");
  const notes = window.document.createElement("div");
  window.document.body.append(container, hidden, pages, notes);
  const hiddenContainer = shallowRef<HTMLElement | null>(null);
  const errors: string[] = [];
  const changes: Document[] = [];
  const holder: { editor: ReturnType<typeof useDocxEditor> | null } = {
    editor: null,
  };
  const app = createApp(
    defineComponent({
      setup() {
        holder.editor = useDocxEditor({
          experimentalSession: "legacy",
          hiddenContainer,
          pagesContainer: shallowRef(pages),
          noteEditorContainer: shallowRef(notes),
          onChange: (changed) => changes.push(changed),
          onError: (error) => errors.push(error.message),
        });
        return () => h("div");
      },
    }),
  );
  app.mount(container);
  // A failed layout pass is recorded, not thrown: collect it with the errors.
  const instrumentation = globalThis.__folioLayoutInstrumentation;
  globalThis.__folioLayoutInstrumentation = {
    onLayoutError: ({ message, reason }) =>
      errors.push(`layout (${reason}): ${message}`),
    onMeasureBlockError: ({ blockKind, message }) =>
      errors.push(`measure (${blockKind}): ${message}`),
  };
  try {
    const adapter = holder.editor ?? panic("Expected mounted Vue editor");
    adapter.loadDocument(document);
    await paint();
    const precomputedPages = adapter.layout.value?.pages.length ?? 0;
    hiddenContainer.value = hidden;
    await paint();
    const view = adapter.editorView.value ?? panic("Expected the body view");
    const mountedPages = adapter.layout.value?.pages.length ?? 0;
    adapter.reLayout();
    const relaidPages = adapter.layout.value?.pages.length ?? 0;
    const read = adapter.readDocument();
    // What `getEditorRef().getDocument()` hands out: the controller's document.
    const paged = adapter.editor.getDocument();
    const copy = adapter.getDocument();
    const saved = await adapter.save();
    typeInto(view.dom, "Z");
    await settle();
    const typed = changes.at(-1) ?? null;
    const typedPages = adapter.layout.value?.pages.length ?? 0;
    const savedAfterEdit = await adapter.save();
    return {
      errors,
      precomputedPages,
      mountedPages,
      relaidPages,
      typedPages,
      read: structuredClone(read),
      paged: structuredClone(paged),
      copy,
      typed,
      saved:
        saved &&
        (await parseShape(new Uint8Array(await saved.arrayBuffer()))).package
          .document,
      savedAfterEdit:
        savedAfterEdit &&
        (await parseShape(new Uint8Array(await savedAfterEdit.arrayBuffer())))
          .package.document,
    };
  } finally {
    globalThis.__folioLayoutInstrumentation = instrumentation;
    app.unmount();
    container.remove();
    hidden.remove();
    pages.remove();
    notes.remove();
  }
};

test.each(DOCUMENT_SHAPES.map((shape) => [shape.id, shape] as const))(
  "the Vue editor reads a frozen %s model as it reads the same model unfrozen",
  async (_id, shape) => {
    const bytes = await shape.build();
    const plain = await driveEditor(await parseShape(bytes));
    const frozen = await driveEditor(deepFreeze(await parseShape(bytes)));

    expect(plain.errors).toEqual([]);
    expect(frozen.errors).toEqual([]);
    expect(plain.precomputedPages).toBeGreaterThan(0);
    expect(plain.typed?.package.document).toBeDefined();
    expect(plain.saved).not.toBeNull();
    expect(frozen).toEqual(plain);
  },
);
