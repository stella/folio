import { GlobalRegistrator } from "@happy-dom/global-registrator";

// ProseMirror's keymap reads the platform at import, before the DOM is
// registered; the editor's page-level shortcuts read it per press.
const KEYMAP_MAC = /Mac|iP(?:hone|[oa]d)/u.test(globalThis.navigator?.platform ?? "");

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import type { EditorView } from "prosemirror-view";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { isMacPlatform } from "@stll/folio-core/managers/editorShortcuts";
import type { HostShortcut } from "@stll/folio-core/managers/editorShortcuts";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

type ModPress = { target: EventTarget; key: string; mac: boolean };

const modPress = ({ target, key, mac }: ModPress): KeyboardEvent => {
  const event = new KeyboardEvent("keydown", {
    key,
    keyCode: key.toUpperCase().charCodeAt(0),
    metaKey: mac,
    ctrlKey: !mac,
    bubbles: true,
    cancelable: true,
  });
  target.dispatchEvent(event);
  return event;
};

/** What the editor reports: its body views and each print it runs. */
const createProbe = () => {
  const views: EditorView[] = [];
  let prints = 0;
  return {
    views,
    prints: () => prints,
    onEditorViewReady: (view: EditorView | null) => {
      if (view !== null) views.push(view);
    },
    onPrint: () => {
      prints += 1;
    },
  };
};

const renderEditor = async (hostShortcuts: readonly HostShortcut[] | undefined) => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  const { views, prints, onEditorViewReady, onPrint } = createProbe();
  await act(async () => {
    root.render(
      <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
        <DocxEditor
          ref={editor}
          document={createEmptyDocument({ initialText: "Hello" })}
          {...(hostShortcuts === undefined ? {} : { hostShortcuts })}
          onEditorViewReady={onEditorViewReady}
          onPrint={onPrint}
          showToolbar={false}
        />
      </IntlProvider>,
    );
  });
  await act(async () => {
    editor.current?.ensureEditorView({ focus: false });
  });
  const view = views.at(-1) ?? panic("The editor did not report its body view");
  return {
    editor: editor.current ?? panic("The editor did not attach its ref"),
    view,
    prints,
    type: async (text: string) => {
      await act(async () => {
        view.dispatch(view.state.tr.insertText(text, view.state.doc.content.size - 1));
      });
    },
    /** A Mod chord inside the body, where the ProseMirror keymap answers it. */
    pressInBody: async (key: string) => {
      await act(async () => {
        modPress({ target: view.dom, key, mac: KEYMAP_MAC });
      });
    },
    /** A Mod chord on the page, where the editor's page-level shortcuts answer it. */
    pressOnPage: async (key: string) => {
      let event: KeyboardEvent | undefined;
      await act(async () => {
        event = modPress({ target: document.body, key, mac: isMacPlatform() });
      });
      return event ?? panic("The key press did not run");
    },
    unmount: async () => {
      await act(async () => root.unmount());
      container.remove();
    },
  };
};

test("the editor answers undo and print by default", async () => {
  const harness = await renderEditor(undefined);
  try {
    await harness.type(" world");
    await harness.pressInBody("z");
    expect(harness.view.state.doc.textContent).toBe("Hello");
    expect((await harness.pressOnPage("p")).defaultPrevented).toBe(true);
    expect(harness.prints()).toBe(1);
  } finally {
    await harness.unmount();
  }
});

test("host-owned shortcuts leave their keys to the host and keep the ref commands", async () => {
  const harness = await renderEditor(["history", "print"]);
  try {
    await harness.type(" world");
    for (const key of ["z", "y"]) {
      await harness.pressInBody(key);
    }
    expect(harness.view.state.doc.textContent).toBe("Hello world");
    expect((await harness.pressOnPage("p")).defaultPrevented).toBe(false);
    expect(harness.prints()).toBe(0);

    await act(async () => {
      expect(harness.editor.undo()).toBe(true);
    });
    expect(harness.view.state.doc.textContent).toBe("Hello");
    await act(async () => {
      expect(harness.editor.redo()).toBe(true);
    });
    expect(harness.view.state.doc.textContent).toBe("Hello world");
  } finally {
    await harness.unmount();
  }
});
