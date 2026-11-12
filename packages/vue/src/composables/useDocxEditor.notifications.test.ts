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
import { afterAll, expect, test } from "bun:test";

const { createApp, defineComponent, h, shallowRef } = await import("vue");

import {
  DOCUMENT_OP_TYPES,
  withCanonicalParagraphIds,
} from "@stll/folio-core/controller/canonicalOperations";
import { createDocx } from "@stll/folio-core/docx/rezip";
import type { Document } from "@stll/folio-core/types/document";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

const { useDocxEditor } = await import("./useDocxEditor");

afterAll(() => GlobalRegistrator.unregister());

/** A deep copy of a whole document, as opposed to a copy of some content. */
const isDocument = (value: unknown): boolean =>
  typeof value === "object" &&
  value !== null &&
  "package" in value &&
  typeof value.package === "object" &&
  value.package !== null &&
  "document" in value.package;

/** Counts whole-document copies made from now until `stop` is called. */
const countDocumentCopies = () => {
  const original = globalThis.structuredClone;
  let copies = 0;
  globalThis.structuredClone = ((
    value: unknown,
    options?: StructuredSerializeOptions,
  ) => {
    if (isDocument(value)) copies += 1;
    return original(value, options);
  }) as typeof structuredClone;
  return {
    copies: () => copies,
    stop: () => {
      globalThis.structuredClone = original;
    },
  };
};

const typeInto = (dom: HTMLElement, data: string) =>
  dom.dispatchEvent(
    new InputEvent("beforeinput", {
      bubbles: true,
      cancelable: true,
      inputType: "insertText",
      data,
    }),
  );

test.each(["note", "body"] as const)(
  "typing in a canonical %s copies the document only for the host, once per notification",
  async (story) => {
    const container = document.createElement("div");
    const hidden = document.createElement("div");
    const pages = document.createElement("div");
    const notes = document.createElement("div");
    document.body.append(container, hidden, pages, notes);
    const changes: Document[] = [];
    const holder: { editor: ReturnType<typeof useDocxEditor> | null } = {
      editor: null,
    };
    const app = createApp(
      defineComponent({
        setup() {
          holder.editor = useDocxEditor({
            hiddenContainer: shallowRef(hidden),
            pagesContainer: shallowRef(pages),
            noteEditorContainer: shallowRef(notes),
            experimentalSession: "canonical",
            onChange: (document) => changes.push(document),
            onError: (error) => panic(error.message),
          });
          return () => h("div");
        },
      }),
    );
    app.mount(container);
    const counter = { stop: () => {} };
    try {
      const adapter = holder.editor ?? panic("Expected mounted Vue editor");
      await adapter.loadBuffer(
        await createDocx(createEmptyDocument({ initialText: "Body" })),
      );
      const before =
        adapter.getDocument() ?? panic("Expected canonical document");
      const paragraph = before.package.document.content.at(0);
      if (paragraph?.type !== "paragraph" || !paragraph.paraId)
        panic("Expected addressed body paragraph");
      expect(
        adapter.editor.applyCanonicalOperations([
          {
            type: DOCUMENT_OP_TYPES.ADD_NOTE,
            at: { story: "main", blockId: paragraph.paraId, offset: 4 },
            note: {
              type: "footnote",
              id: 1,
              content: withCanonicalParagraphIds(
                [{ type: "paragraph", content: [] }],
                before,
              ),
            },
          },
        ]),
      ).toBe(true);
      const view =
        story === "note"
          ? (adapter.openNoteStory({ kind: "footnote", noteId: 1 }),
            adapter.getActiveNoteView())
          : adapter.editorView.value;
      if (!view) panic(`Expected the ${story} view`);
      await new Promise((resolve) => setTimeout(resolve, 600));
      const emitted: Document[] = [];
      const off = adapter.editor.on("docChange", (document) =>
        emitted.push(document),
      );
      changes.length = 0;

      const typing = countDocumentCopies();
      counter.stop = typing.stop;
      typeInto(view.dom, "N");
      // The keystroke itself copies nothing: the editor reads the session's
      // document where it stands.
      expect(typing.copies()).toBe(0);
      await new Promise((resolve) => setTimeout(resolve, 600));
      typing.stop();
      off();

      expect(view.state.doc.textContent).toContain("N");
      expect(changes).toHaveLength(1);
      expect(emitted).toHaveLength(1);
      // Both listeners get one copy, the same one, showing the edit.
      expect(typing.copies()).toBe(1);
      expect(emitted[0]).toBe(changes[0]);
      const notified = changes[0] ?? panic("Expected a notified document");
      expect(notified).toEqual(
        adapter.getDocument() ?? panic("Expected canonical document"),
      );
      // And the host owns it: emptying it leaves the editor's document alone.
      notified.package.document.content.length = 0;
      expect(
        adapter.getDocument()?.package.document.content.length,
      ).toBeGreaterThan(0);
    } finally {
      counter.stop();
      app.unmount();
      container.remove();
      hidden.remove();
      pages.remove();
      notes.remove();
    }
  },
);
