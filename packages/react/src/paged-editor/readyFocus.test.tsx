import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

import { panic } from "better-result";
import { afterAll, expect, test } from "bun:test";
import { TextSelection } from "prosemirror-state";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import { PagedEditor, type PagedEditorRef } from "./PagedEditor";

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

/** Animation frames held until `flush`, so a test decides when they run. */
const holdAnimationFrames = () => {
  const original = globalThis.requestAnimationFrame;
  const held: FrameRequestCallback[] = [];
  globalThis.requestAnimationFrame = (callback) => {
    held.push(callback);
    return held.length;
  };
  return {
    flush: () => {
      for (const callback of held.splice(0)) callback(performance.now());
    },
    release: () => {
      globalThis.requestAnimationFrame = original;
    },
  };
};

/**
 * `selectionchange` held back from the editor until `deliver`, as when the
 * browser reports a caret move after the next animation frame.
 */
const holdSelectionChanges = () => {
  let held = false;
  const swallow = (event: Event) => {
    if (!held) return;
    event.stopImmediatePropagation();
  };
  globalThis.addEventListener("selectionchange", swallow, true);
  return {
    hold: () => {
      held = true;
    },
    deliver: () => {
      held = false;
      document.dispatchEvent(new Event("selectionchange"));
    },
    release: () => {
      globalThis.removeEventListener("selectionchange", swallow, true);
    },
  };
};

const documentIO = {
  getDocx: async () => new ArrayBuffer(0),
  loadDocument: () => undefined,
  loadDocx: async () => undefined,
};

// A click that creates the hidden view focuses it at once and finishes its
// setup a frame later. A key pressed in between moves the browser's caret,
// which ProseMirror reads from a later `selectionchange`; the frame must not
// put the editor's older selection back over it.
test("the frame after the view is created keeps a caret a key moved", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<PagedEditorRef>();
  const frames = holdAnimationFrames();
  const selectionChanges = holdSelectionChanges();
  try {
    await act(async () =>
      root.render(
        <IntlProvider
          locale="en"
          timeZone="UTC"
          messages={getFolioMessages("en")}
        >
          <PagedEditor
            accessibleName="Document editor"
            ref={editor}
            document={createEmptyDocument({
              initialText: "Please remit funds.",
            })}
            documentIdentity="ready-focus"
            markupView="all-markup"
            documentIO={documentIO}
          />
        </IntlProvider>,
      ),
    );
    await act(async () => editor.current?.ensureView());
    const view =
      editor.current?.getEditor()?.getView() ?? panic("Expected a hidden view");
    // The click's caret, after "Please re".
    const clicked = 1 + "Please re".length;
    await act(async () => {
      view.focus();
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(view.state.doc, clicked),
        ),
      );
    });
    // End: the browser moves its caret to the line end.
    const lineEnd = 1 + "Please remit funds.".length;
    const moved = view.domAtPos(lineEnd);
    const native = document.getSelection() ?? panic("Expected a selection");
    selectionChanges.hold();
    native.collapse(moved.node, moved.offset);

    await act(async () => frames.flush());
    await act(async () => selectionChanges.deliver());

    const after = document.getSelection() ?? panic("Expected a selection");
    const anchor =
      after.anchorNode ?? panic("Expected the caret in the document");
    expect(view.posAtDOM(anchor, after.anchorOffset)).toBe(lineEnd);
    expect(view.state.selection.head).toBe(lineEnd);
  } finally {
    selectionChanges.release();
    frames.release();
    await act(async () => root.unmount());
    container.remove();
  }
});
