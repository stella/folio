import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

import { panic } from "better-result";
import { afterAll, expect, test } from "bun:test";
import { Plugin, TextSelection } from "prosemirror-state";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import {
  getMeasureProvider,
  type MeasureProvider,
  setMeasureProvider,
} from "@stll/folio-core/layout-engine/measure/measureProvider";
import type {
  FlowBlock,
  Layout,
  Measure,
} from "@stll/folio-core/layout-engine/types";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import {
  HiddenProseMirror,
  type HiddenProseMirrorRef,
} from "./HiddenProseMirror";
import { PagedEditor, type PagedEditorRef } from "./PagedEditor";
import { useSelectionOverlay } from "./SelectionOverlay";

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const createTransactionCallbacks = (
  calls: string[],
  callbackLabel: string,
) => ({
  onTransaction: () => {
    calls.push(callbackLabel);
  },
  onReadOnlyEditAttempt: () => {
    calls.push(`blocked-${callbackLabel}`);
  },
});

const createDocumentIO = (buffer: ArrayBuffer) => ({
  getDocx: async () => buffer,
  loadDocument: () => undefined,
  loadDocx: async () => undefined,
});

const createSeedProbe = () => {
  let initializations = 0;
  return {
    plugins: [
      new Plugin({
        state: {
          init: () => {
            initializations++;
            return null;
          },
          apply: () => null,
        },
      }),
    ],
    getInitializations: () => initializations,
  };
};

const createSelectionMeasureProvider = (installed: MeasureProvider) =>
  ({
    ...installed,
    measureTextWidth: (text) => text.length * 7,
    measureRun: (text) => ({
      width: text.length * 7,
      charWidths: Array.from(text, () => 7),
      metrics: {
        fontSize: 11,
        ascent: 12,
        descent: 4,
        fontBoxAscent: 12,
        fontBoxDescent: 4,
        lineHeight: 16,
        fontFamily: "Calibri",
        singleLineRatio: 1,
      },
    }),
  }) satisfies MeasureProvider;

const createSelectionFixture = () => {
  const blocks = [
    {
      kind: "paragraph",
      id: "p1",
      pmStart: 0,
      pmEnd: 5,
      runs: [
        {
          kind: "text",
          text: "abc",
          fontFamily: "Calibri",
          fontSize: 11,
          pmStart: 1,
          pmEnd: 4,
        },
      ],
    },
  ] satisfies FlowBlock[];
  const measures = [
    {
      kind: "paragraph",
      width: 21,
      height: 16,
      lines: [
        {
          fromRun: 0,
          toRun: 0,
          fromChar: 0,
          toChar: 3,
          width: 21,
          lineHeight: 16,
          ascent: 12,
          descent: 4,
        },
      ],
    },
  ] satisfies Measure[];
  const fixture = {
    pageSize: { w: 600, h: 800 },
    pageGap: 0,
    pages: [
      {
        number: 1,
        size: { w: 600, h: 800 },
        margins: { top: 0, right: 0, bottom: 0, left: 0 },
        fragments: [
          {
            kind: "paragraph",
            blockId: "p1",
            x: 10,
            y: 20,
            width: 500,
            height: 16,
            fromLine: 0,
            toLine: 1,
          },
        ],
      },
    ],
  } satisfies Layout;
  return {
    blocks,
    measures,
    fixture,
    collapsedSelection: { from: 1, to: 1 },
    rangeSelection: { from: 1, to: 4 },
  };
};

test("persistent hidden view reads committed callbacks and read-only inputs", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<HiddenProseMirrorRef>();
  const documentModel = createEmptyDocument({ initialText: "Initial" });
  const calls: string[] = [];
  const render = async (readOnly: boolean, callbackLabel: string) => {
    const { onTransaction, onReadOnlyEditAttempt } = createTransactionCallbacks(
      calls,
      callbackLabel,
    );
    await act(async () =>
      root.render(
        <HiddenProseMirror
          ref={editor}
          document={documentModel}
          documentIdentity="initial"
          readOnly={readOnly}
          onTransaction={onTransaction}
          onReadOnlyEditAttempt={onReadOnlyEditAttempt}
        />,
      ),
    );
  };
  try {
    await render(false, "first");
    await act(async () => editor.current?.ensureView());
    const view = editor.current?.getView();
    expect(view).toBeDefined();
    if (!view) panic("Expected a hidden view");
    await act(async () => view.dispatch(view.state.tr.insertText("!", 1)));
    expect(calls).toEqual(["first"]);
    await render(true, "second");
    expect(editor.current?.getView()).toBe(view);
    await act(async () => view.dispatch(view.state.tr.insertText("?", 1)));
    expect(calls).toEqual(["first", "blocked-second"]);
    await render(false, "third");
    expect(editor.current?.getView()).toBe(view);
    await act(async () => view.dispatch(view.state.tr.insertText("+", 1)));
    expect(calls).toEqual(["first", "blocked-second", "third"]);
    expect(view.state.doc.textContent).toBe("+!Initial");
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});

test("selection geometry follows collapsed, range, and cleared inputs without stale effects", async () => {
  const container = document.createElement("div");
  const root = createRoot(container);
  const installedProvider = getMeasureProvider();
  setMeasureProvider(createSelectionMeasureProvider(installedProvider));
  const { blocks, measures, fixture, collapsedSelection, rangeSelection } =
    createSelectionFixture();
  const Probe = ({
    selection,
    layout,
  }: {
    selection: { from: number; to: number } | null;
    layout: Layout | null;
  }) => {
    const geometry = useSelectionOverlay(selection, layout, blocks, measures);
    return (
      <output
        data-caret={geometry.caretPosition === null ? "absent" : "present"}
        data-rects={geometry.selectionRects.length}
        data-selection={JSON.stringify(geometry.selectionRects)}
      />
    );
  };
  try {
    await act(async () =>
      root.render(<Probe selection={collapsedSelection} layout={fixture} />),
    );
    expect(container.querySelector("output")?.getAttribute("data-caret")).toBe(
      "present",
    );
    expect(container.querySelector("output")?.getAttribute("data-rects")).toBe(
      "0",
    );
    await act(async () =>
      root.render(<Probe selection={rangeSelection} layout={fixture} />),
    );
    expect(container.querySelector("output")?.getAttribute("data-caret")).toBe(
      "absent",
    );
    expect(container.querySelector("output")?.getAttribute("data-rects")).toBe(
      "1",
    );
    expect(
      JSON.parse(
        container.querySelector("output")?.getAttribute("data-selection") ??
          "null",
      ),
    ).toEqual([{ x: 10, y: 20, width: 21, height: 16, pageIndex: 0 }]);
    for (const input of [
      { selection: null, layout: fixture },
      { selection: { from: 1, to: 1 }, layout: null },
    ]) {
      await act(async () => root.render(<Probe {...input} />));
      expect(
        container.querySelector("output")?.getAttribute("data-caret"),
      ).toBe("absent");
      expect(
        container.querySelector("output")?.getAttribute("data-rects"),
      ).toBe("0");
    }
  } finally {
    await act(async () => root.unmount());
    setMeasureProvider(installedProvider);
  }
});

test("persistent paged controller reads document I/O from the latest commit", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<PagedEditorRef>();
  const messages = getFolioMessages("en");
  const render = async (buffer: ArrayBuffer) => {
    const documentIO = createDocumentIO(buffer);
    await act(async () =>
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={messages}>
          <PagedEditor
            ref={editor}
            document={null}
            documentIdentity="committed-document-io"
            markupView="all-markup"
            documentIO={documentIO}
          />
        </IntlProvider>,
      ),
    );
  };
  try {
    const first = new ArrayBuffer(1);
    await render(first);
    const controller =
      editor.current?.getEditor() ?? panic("Expected an editor controller");
    expect(await controller.getDocx()).toBe(first);
    const replacement = new ArrayBuffer(2);
    await render(replacement);
    expect(editor.current?.getEditor()).toBe(controller);
    expect(await controller.getDocx()).toBe(replacement);
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});

test("read-only pages route selection keys but preserve native scrolling and refuse edits", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<PagedEditorRef>();
  const messages = getFolioMessages("en");
  const calls: string[] = [];
  const { onReadOnlyEditAttempt } = createTransactionCallbacks(
    calls,
    "readonly-keyboard",
  );
  await act(async () =>
    root.render(
      <IntlProvider locale="en" timeZone="UTC" messages={messages}>
        <PagedEditor
          ref={editor}
          document={createEmptyDocument({ initialText: "Read only text" })}
          documentIdentity="readonly-keyboard-selection"
          markupView="all-markup"
          documentIO={createDocumentIO(new ArrayBuffer(0))}
          readOnly
          onReadOnlyEditAttempt={onReadOnlyEditAttempt}
        />
      </IntlProvider>,
    ),
  );
  try {
    await act(async () => editor.current?.ensureView({ focus: false }));
    const view =
      editor.current?.getEditor()?.getView() ?? panic("Expected a hidden view");
    const pageSurface = container.querySelector<HTMLElement>(".paged-editor");
    if (!pageSurface) panic("Expected the page surface");
    const dispatchKey = (key: string, init: KeyboardEventInit = {}) => {
      const event = new KeyboardEvent("keydown", {
        key,
        bubbles: true,
        cancelable: true,
        ...init,
      });
      pageSurface.dispatchEvent(event);
      return event;
    };

    await act(async () => {
      view.dispatch(
        view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)),
      );
      dispatchKey("a", { ctrlKey: true });
    });
    expect(view.state.selection.from).toBe(1);
    expect(view.state.selection.to).toBe(view.state.doc.content.size - 1);

    await act(async () => {
      view.dispatch(
        view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)),
      );
      const event = dispatchKey("ArrowRight", { shiftKey: true });
      expect(event.defaultPrevented).toBe(true);
    });
    expect([view.state.selection.from, view.state.selection.to]).toEqual([
      1, 2,
    ]);

    const before = view.state.doc;
    await act(async () => {
      const event = dispatchKey("b");
      expect(event.defaultPrevented).toBe(true);
    });
    expect(calls).toEqual(["blocked-readonly-keyboard"]);
    expect(view.state.doc).toBe(before);

    await act(async () => {
      const event = dispatchKey("ArrowLeft");
      expect(event.defaultPrevented).toBe(false);
    });
    expect(view.state.doc).toBe(before);
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});

test("deferred hidden view reuses the latest committed pre-view document seed", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<PagedEditorRef>();
  const messages = getFolioMessages("en");
  const probe = createSeedProbe();
  const documentIO = createDocumentIO(new ArrayBuffer(0));
  const render = async (
    documentModel: ReturnType<typeof createEmptyDocument>,
    identity: string,
  ) => {
    await act(async () =>
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={messages}>
          <PagedEditor
            ref={editor}
            document={documentModel}
            documentIdentity={identity}
            markupView="all-markup"
            documentIO={documentIO}
            externalPlugins={probe.plugins}
          />
        </IntlProvider>,
      ),
    );
  };
  try {
    await render(createEmptyDocument({ initialText: "First" }), "first");
    const controller =
      editor.current?.getEditor() ?? panic("Expected an editor controller");
    expect(controller.getView()).toBeNull();
    await render(
      createEmptyDocument({ initialText: "Replacement" }),
      "replacement",
    );
    expect(editor.current?.getEditor()).toBe(controller);
    expect(controller.getView()).toBeNull();
    const initializedBeforeRequest = probe.getInitializations();
    expect(initializedBeforeRequest).toBeGreaterThan(0);
    await act(async () => editor.current?.ensureView({ focus: false }));
    const view =
      controller.getView() ?? panic("Expected the requested hidden view");
    expect(view.state.doc.textContent).toBe("Replacement");
    expect(probe.getInitializations()).toBe(initializedBeforeRequest);
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});
