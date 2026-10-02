import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { Plugin, PluginKey } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import { PagedEditor } from "../paged-editor/PagedEditor";
import type { PagedEditorRef } from "../paged-editor/PagedEditor";
import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const createHostPluginFixture = () => {
  const key = new PluginKey("host-append-text");
  const plugin = new Plugin({
    key,
    appendTransaction(transactions, _oldState, state) {
      return transactions.some((transaction) => transaction.getMeta(key) === true)
        ? state.tr.insertText("!", state.doc.content.size - 1)
        : null;
    },
  });
  const plugins = [plugin];
  const views: EditorView[] = [];
  const onEditorViewReady = (view: EditorView | null) => {
    if (view !== null) {
      views.push(view);
    }
  };
  return { key, plugin, plugins, views, onEditorViewReady };
};

test("host plugins process transactions after replacing the body document", async () => {
  const { key, plugin, plugins, views, onEditorViewReady } = createHostPluginFixture();
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  const renderDocument = async (text: string) => {
    const document = createEmptyDocument({ initialText: text });
    await act(async () => {
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
          <DocxEditor
            ref={editor}
            document={document}
            plugins={plugins}
            onEditorViewReady={onEditorViewReady}
            showToolbar={false}
          />
        </IntlProvider>,
      );
    });
  };
  try {
    for (const text of ["First document", "Replacement document"]) {
      await renderDocument(text);
      await act(async () => {
        editor.current?.ensureEditorView({ focus: false });
      });
      const view = views.at(-1) ?? panic("The editor did not report its body view");
      expect(view.state.doc.textContent).toBe(text);
      expect(view.state.plugins).toContain(plugin);
      await act(async () => {
        view.dispatch(view.state.tr.setMeta(key, true));
      });
      expect(view.state.doc.textContent).toBe(`${text}!`);
    }
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});

test("replacing the document resets the scroll position", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const renderDocument = async (text: string) => {
    await act(async () => {
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
          <DocxEditor document={createEmptyDocument({ initialText: text })} showToolbar={false} />
        </IntlProvider>,
      );
    });
  };
  try {
    await renderDocument("First document");
    const scrollContainer = container.querySelector("[data-folio-scroll]");
    if (!(scrollContainer instanceof HTMLElement)) {
      panic("The editor did not render its scroll container");
    }
    scrollContainer.scrollTop = 4321;
    scrollContainer.scrollLeft = 17;

    await renderDocument("Replacement document");

    expect(scrollContainer.scrollTop).toBe(0);
    expect(scrollContainer.scrollLeft).toBe(0);
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});

const createReadyScrollFixture = () => {
  const editor = createRef<DocxEditorRef>();
  const onEditorViewReady = (view: EditorView | null) => {
    if (view) {
      const scrollRoot = editor.current?.getScrollRoot();
      if (scrollRoot) scrollRoot.scrollTop = 321;
    }
  };
  return { editor, onEditorViewReady };
};

const documentIO = {
  getDocx: async () => null,
  loadDocument: () => {},
  loadDocx: async () => {},
};

test("host view-ready navigation wins over a pending initial scroll offset", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const { editor, onEditorViewReady } = createReadyScrollFixture();
  const frames = new Map<number, FrameRequestCallback>();
  let frameId = 0;
  const previousRequest = globalThis.requestAnimationFrame;
  const previousCancel = globalThis.cancelAnimationFrame;
  globalThis.requestAnimationFrame = (callback) => {
    frameId += 1;
    frames.set(frameId, callback);
    return frameId;
  };
  globalThis.cancelAnimationFrame = (id) => {
    frames.delete(id);
  };
  try {
    await act(async () => {
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
          <DocxEditor
            ref={editor}
            document={createEmptyDocument({ initialText: "Initial scroll precedence" })}
            showToolbar={false}
            initialScrollTop={42}
            onEditorViewReady={onEditorViewReady}
          />
        </IntlProvider>,
      );
    });
    await act(async () => {
      editor.current?.ensureEditorView({ focus: false });
    });
    const scrollRoot = editor.current?.getScrollRoot() ?? panic("The editor has no scroll root");
    expect(scrollRoot.scrollTop).toBe(321);
    const scheduled = [...frames.values()];
    frames.clear();
    await act(async () => {
      for (const callback of scheduled) callback(performance.now());
    });
    expect(scrollRoot.scrollTop).toBe(321);
  } finally {
    await act(async () => root.unmount());
    globalThis.requestAnimationFrame = previousRequest;
    globalThis.cancelAnimationFrame = previousCancel;
    container.remove();
  }
});

for (const existingMarker of [undefined, "host-owned"]) {
  test(`external paged-editor scroll roots preserve marker ownership (${existingMarker ?? "unmarked"})`, async () => {
    const scrollRoot = document.createElement("div");
    if (existingMarker !== undefined) scrollRoot.setAttribute("data-folio-scroll", existingMarker);
    const container = document.createElement("div");
    scrollRoot.append(container);
    document.body.append(scrollRoot);
    const root = createRoot(container);
    const editor = createRef<PagedEditorRef>();
    const scrollContainerRef = createRef<HTMLDivElement>();
    scrollContainerRef.current = scrollRoot;
    try {
      await act(async () => {
        root.render(
          <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
            <PagedEditor
              ref={editor}
              document={null}
              documentIdentity="scroll-root-lifecycle"
              documentIO={documentIO}
              markupView="all-markup"
              scrollContainerRef={scrollContainerRef}
            />
          </IntlProvider>,
        );
      });
      expect(editor.current?.getScrollRoot()).toBe(scrollRoot);
      expect(scrollRoot.getAttribute("data-folio-scroll")).toBe(existingMarker ?? "");
      await act(async () => root.unmount());
      expect(scrollRoot.getAttribute("data-folio-scroll")).toBe(existingMarker ?? null);
    } finally {
      scrollRoot.remove();
    }
  });
}
