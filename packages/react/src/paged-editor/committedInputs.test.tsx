import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { HiddenProseMirror, type HiddenProseMirrorRef } from "./HiddenProseMirror";
import type { FlowBlock, Layout, Measure } from "@stll/folio-core/layout-engine/types";
import { useSelectionOverlay } from "./SelectionOverlay";
import { PagedEditor, type PagedEditorRef } from "./PagedEditor";
import { IntlProvider } from "use-intl";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const createTransactionCallbacks = (calls: string[], callbackLabel: string) => ({
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

const createSelectionFixture = () => {
  const blocks = [
    {
      kind: "paragraph",
      id: "p1",
      pmStart: 0,
      pmEnd: 5,
      runs: [
        { kind: "text", text: "abc", fontFamily: "Calibri", fontSize: 11, pmStart: 1, pmEnd: 4 },
      ],
    },
  ] satisfies FlowBlock[];
  const measures = [
    {
      kind: "paragraph",
      width: 35,
      height: 16,
      lines: [
        {
          fromRun: 0,
          toRun: 0,
          fromChar: 0,
          toChar: 3,
          width: 35,
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
      />
    );
  };
  try {
    await act(async () => root.render(<Probe selection={collapsedSelection} layout={fixture} />));
    expect(container.querySelector("output")?.getAttribute("data-caret")).toBe("present");
    expect(container.querySelector("output")?.getAttribute("data-rects")).toBe("0");
    await act(async () => root.render(<Probe selection={rangeSelection} layout={fixture} />));
    expect(container.querySelector("output")?.getAttribute("data-caret")).toBe("absent");
    expect(container.querySelector("output")?.getAttribute("data-rects")).toBe("1");
    for (const input of [
      { selection: null, layout: fixture },
      { selection: { from: 1, to: 1 }, layout: null },
    ]) {
      await act(async () => root.render(<Probe {...input} />));
      expect(container.querySelector("output")?.getAttribute("data-caret")).toBe("absent");
      expect(container.querySelector("output")?.getAttribute("data-rects")).toBe("0");
    }
  } finally {
    await act(async () => root.unmount());
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
    const controller = editor.current?.getEditor() ?? panic("Expected an editor controller");
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
