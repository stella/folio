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

const { createApp, defineComponent, h, isProxy, reactive, shallowRef } =
  await import("vue");

import {
  DOCUMENT_OP_TYPES,
  withCanonicalParagraphIds,
} from "@stll/folio-core/controller/canonicalOperations";
import { cloneDocumentWithParagraphPropertySources } from "@stll/folio-core/docx/document-clone";
import { createDocx } from "@stll/folio-core/docx/rezip";
import type { Document } from "@stll/folio-core/types/document";
import { readModel } from "@stll/folio-core/types/documentModel";
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
  "typing in a canonical %s copies the document only for the host, once per settled change",
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
      typeInto(view.dom, "O");
      typeInto(view.dom, "P");
      // The keystrokes themselves copy nothing: the editor reads the
      // session's document where it stands, and the host hears of the edits
      // once they settle.
      expect(typing.copies()).toBe(0);
      expect(changes).toHaveLength(0);
      await new Promise((resolve) => setTimeout(resolve, 600));
      typing.stop();
      off();

      expect(view.state.doc.textContent).toContain("NOP");
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

test("a reactive document a host loads is held raw, so the host's copies can clone it", () => {
  const container = document.createElement("div");
  document.body.append(container);
  const holder: { editor: ReturnType<typeof useDocxEditor> | null } = {
    editor: null,
  };
  const app = createApp(
    defineComponent({
      setup() {
        holder.editor = useDocxEditor({
          hiddenContainer: shallowRef(null),
          pagesContainer: shallowRef(null),
          onError: (error) => panic(error.message),
        });
        return () => h("div");
      },
    }),
  );
  app.mount(container);
  try {
    const adapter = holder.editor ?? panic("Expected mounted Vue editor");
    const hosted = reactive(createEmptyDocument({ initialText: "Body" }));
    adapter.loadDocument(hosted);
    const model = adapter.modelDocument.value ?? panic("Expected a model");
    expect(isProxy(model)).toBe(false);
    expect(() =>
      cloneDocumentWithParagraphPropertySources(model),
    ).not.toThrow();
  } finally {
    app.unmount();
    container.remove();
  }
});

/** Mounts a Vue editor in the default session, loaded with `text`. */
const mountDefaultEditor = async (text: string) => {
  const container = document.createElement("div");
  const hidden = document.createElement("div");
  const pages = document.createElement("div");
  document.body.append(container, hidden, pages);
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
          onChange: (changed) => changes.push(changed),
          onError: (error) => panic(error.message),
        });
        return () => h("div");
      },
    }),
  );
  app.mount(container);
  const adapter = holder.editor ?? panic("Expected mounted Vue editor");
  await adapter.loadBuffer(
    await createDocx(createEmptyDocument({ initialText: text })),
  );
  await new Promise((resolve) => setTimeout(resolve, 600));
  changes.length = 0;
  return {
    adapter,
    changes,
    unmount: () => {
      app.unmount();
      container.remove();
      hidden.remove();
      pages.remove();
    },
  };
};

test("typing in the default session's body hands the host one copy per settled change, never the model", async () => {
  const mounted = await mountDefaultEditor("Body");
  const counter = { stop: () => {} };
  try {
    const { adapter, changes } = mounted;
    const view = adapter.editorView.value ?? panic("Expected the body view");
    const typing = countDocumentCopies();
    counter.stop = typing.stop;
    typeInto(view.dom, "N");
    typeInto(view.dom, "O");
    typeInto(view.dom, "P");
    expect(typing.copies()).toBe(0);
    await new Promise((resolve) => setTimeout(resolve, 600));
    typing.stop();

    expect(changes).toHaveLength(1);
    expect(typing.copies()).toBe(1);
    const notified = changes[0] ?? panic("Expected a notified document");
    const model = adapter.modelDocument.value ?? panic("Expected a model");
    expect(notified).toEqual(readModel(model));
    // The host owns its copy: emptying it leaves the model alone.
    notified.package.document.content.length = 0;
    expect(model.package.document.content.length).toBeGreaterThan(0);
  } finally {
    counter.stop();
    mounted.unmount();
  }
});

test("the editor's readers read the document without a copy; the host's getDocument makes one", async () => {
  const mounted = await mountDefaultEditor("Body");
  const counter = { stop: () => {} };
  try {
    const { adapter } = mounted;
    const reading = countDocumentCopies();
    counter.stop = reading.stop;
    const read = adapter.readDocument() ?? panic("Expected a document");
    adapter.readDocument();
    adapter.readDocument();
    expect(reading.copies()).toBe(0);
    const copy = adapter.getDocument() ?? panic("Expected a document");
    reading.stop();
    expect(reading.copies()).toBe(1);
    expect(copy).toEqual(readModel(read));
  } finally {
    counter.stop();
    mounted.unmount();
  }
});

test("a document holding a reactive proxy deeper down is held as raw data", () => {
  const container = document.createElement("div");
  document.body.append(container);
  const holder: { editor: ReturnType<typeof useDocxEditor> | null } = {
    editor: null,
  };
  const app = createApp(
    defineComponent({
      setup() {
        holder.editor = useDocxEditor({
          hiddenContainer: shallowRef(null),
          pagesContainer: shallowRef(null),
          onError: (error) => panic(error.message),
        });
        return () => h("div");
      },
    }),
  );
  app.mount(container);
  try {
    const adapter = holder.editor ?? panic("Expected mounted Vue editor");
    const plain = createEmptyDocument({ initialText: "Body" });
    // A spread of a reactive document is a plain object holding proxies.
    const hosted = { ...reactive(plain) };
    expect(isProxy(hosted)).toBe(false);
    expect(isProxy(hosted.package)).toBe(true);
    adapter.loadDocument(hosted);
    const model = adapter.modelDocument.value ?? panic("Expected a model");
    expect(isProxy(model.package)).toBe(false);
    expect(() => adapter.getDocument()).not.toThrow();
    expect(adapter.getDocument()?.package.document.content).toEqual(
      plain.package.document.content,
    );
    // The host's own object is left as it was.
    expect(isProxy(hosted.package)).toBe(true);
  } finally {
    app.unmount();
    container.remove();
  }
});

test("the controller's getDocument in the default session is read-only: a host writes its own copy, and the editor's model stays", async () => {
  const mounted = await mountDefaultEditor("Body");
  try {
    const { adapter } = mounted;
    const model = adapter.modelDocument.value ?? panic("Expected a model");
    const margin = () =>
      model.package.document.finalSectionProperties?.marginLeft;
    const before = margin();
    expect(before).toBeNumber();

    // The document shares the editor's model, so it is handed out read-only
    // (writing it is a type error): a host that writes copies it first.
    const handed =
      adapter.editor.getDocument() ??
      panic("Expected the controller's document");
    const owned = cloneDocumentWithParagraphPropertySources(handed);
    const sections =
      owned.package.document.finalSectionProperties ??
      panic("Expected section properties");
    sections.marginLeft = 4321;

    expect(margin()).toBe(before);
    expect(
      adapter.getDocument()?.package.document.finalSectionProperties
        ?.marginLeft,
    ).toBe(before);
  } finally {
    mounted.unmount();
  }
});
