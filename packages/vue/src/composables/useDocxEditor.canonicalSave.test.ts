import { GlobalRegistrator } from "@happy-dom/global-registrator";

// PM keymaps capture the host platform at import time. Keep the DOM fixture on
// that same platform when another test has already imported the editor schema.
const hostPlatform = typeof navigator === "undefined" ? "" : navigator.platform;
GlobalRegistrator.register();
Object.defineProperty(navigator, "platform", { configurable: true, value: hostPlatform });

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { TextSelection } from "prosemirror-state";
import { closeHistory } from "prosemirror-history";

const { createApp, defineComponent, h, shallowRef } = await import("vue");

import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { CanonicalDocxInputError } from "@stll/folio-core/docx/canonicalSessionInput";
import { createEmptyHeaderFooter } from "@stll/folio-core/utils/headerFooter";

const { isMacPlatform } = await import("@stll/folio-core/managers/editorShortcuts");
import {
  createCanonicalHeaderFooterOperation,
  createCanonicalSectionPropertiesOperation,
  DOCUMENT_OP_TYPES,
  removeCanonicalHeaderFooterOperations,
  withCanonicalParagraphIds,
} from "@stll/folio-core/controller/canonicalOperations";

const { usePageSetupControls } = await import("./usePageSetupControls");
import { reviewDifferences } from "../../../../test/reviewDifferences";
import { canonicalReviewBlocks } from "../../../../test/reviewProjection";

const { useDocxEditor } = await import("./useDocxEditor");
const { useKeyboardShortcuts } = await import("./useKeyboardShortcuts");

// The save oracle exercises the composable's real hidden manager and serialization,
// rather than rebuilding the expected document from its PM projection.
afterAll(() => GlobalRegistrator.unregister());

test.each([
  {
    kind: "encrypted",
    bytes: new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
    message: "Password-protected documents are unavailable in the experimental canonical session.",
  },
  {
    kind: "malformed",
    bytes: new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
    message: "Failed to normalize paragraph IDs: Failed to parse DOCX archive",
  },
])("canonical $kind loads preserve the typed input error", async ({ bytes, message }) => {
  const container = document.createElement("div");
  document.body.append(container);
  const errors: Error[] = [];
  const holder: { editor: ReturnType<typeof useDocxEditor> | null } = { editor: null };
  const app = createApp(
    defineComponent({
      setup() {
        holder.editor = useDocxEditor({
          hiddenContainer: shallowRef(null),
          pagesContainer: shallowRef(null),
          experimentalSession: "canonical",
          onError: (error) => errors.push(error),
        });
        return () => h("div");
      },
    }),
  );
  app.mount(container);
  try {
    const editor = holder.editor ?? panic("Expected mounted Vue editor");
    await editor.loadBuffer(bytes);
    expect(errors).toHaveLength(1);
    expect(errors.at(0)).toBeInstanceOf(CanonicalDocxInputError);
    expect(errors.at(0)?.message).toBe(message);
    expect(editor.parseError.value).toBe(message);
    expect(editor.isReady.value).toBe(false);
    expect(editor.getDocument()).toBeNull();
  } finally {
    app.unmount();
    container.remove();
  }
});
test.each(["canonical", "default"] as const)(
  "%s history keys have one editor owner and no document-level owner",
  async (session) => {
    const container = document.createElement("div");
    const hidden = document.createElement("div");
    const pages = document.createElement("div");
    document.body.append(container, hidden, pages);
    const holder: { editor: ReturnType<typeof useDocxEditor> | null } = { editor: null };
    const app = createApp(
      defineComponent({
        setup() {
          holder.editor = useDocxEditor({
            hiddenContainer: shallowRef(hidden),
            pagesContainer: shallowRef(pages),
            ...(session === "canonical" ? { experimentalSession: "canonical" as const } : {}),
          });
          useKeyboardShortcuts({
            showFindReplace: shallowRef(false),
            showHyperlink: shallowRef(false),
            handleZoomKeyDown: () => {},
            scope: () => "document",
            roots: [shallowRef(container)],
          });
          return () => h("div");
        },
      }),
    );
    app.mount(container);
    try {
      const editor = holder.editor ?? panic("Expected mounted Vue editor");
      await editor.loadBuffer(await createDocx(createEmptyDocument({ initialText: "Start" })));
      const view = editor.editorView.value ?? panic("Expected body editor view");
      const insert = (text: string) => {
        view.dispatch(
          closeHistory(
            view.state.tr.setSelection(
              TextSelection.create(view.state.doc, view.state.doc.content.size - 1),
            ),
          ),
        );
        if (session === "canonical") {
          view.dom.dispatchEvent(
            new InputEvent("beforeinput", {
              bubbles: true,
              cancelable: true,
              inputType: "insertText",
              data: text,
            }),
          );
        } else {
          view.dispatch(view.state.tr.insertText(text));
        }
      };
      insert("1");
      insert("2");
      const key = (chord: string) =>
        new KeyboardEvent("keydown", {
          bubbles: true,
          cancelable: true,
          ctrlKey: !isMacPlatform(),
          metaKey: isMacPlatform(),
          key: chord,
        });
      for (const chord of ["z", "y"]) {
        const event = key(chord);
        document.dispatchEvent(event);
        expect(event.defaultPrevented).toBe(false);
        expect(view.state.doc.textContent).toBe("Start12");
      }
      const undo = key("z");
      view.dom.dispatchEvent(undo);
      expect(undo.defaultPrevented).toBe(true);
      expect(view.state.doc.textContent).toBe("Start1");
      const redo = key("y");
      view.dom.dispatchEvent(redo);
      expect(redo.defaultPrevented).toBe(true);
      expect(view.state.doc.textContent).toBe("Start12");
    } finally {
      app.unmount();
      container.remove();
      hidden.remove();
      pages.remove();
    }
  },
);

test("canonical edits save and reopen the canonical text and paragraph identity", async () => {
  const container = document.createElement("div");
  const hidden = document.createElement("div");
  const pages = document.createElement("div");
  document.body.append(container, hidden, pages);
  const errors: Error[] = [];
  let hostChanges = 0;
  let hostInputs = 0;
  const holder: { editor: ReturnType<typeof import("./useDocxEditor").useDocxEditor> | null } = {
    editor: null,
  };
  const app = createApp(
    defineComponent({
      setup() {
        holder.editor = useDocxEditor({
          hiddenContainer: shallowRef(hidden),
          pagesContainer: shallowRef(pages),
          experimentalSession: "canonical",
          onError: (error) => errors.push(error),
          onChange: (document) => {
            hostChanges++;
            document.package.document.content = [];
          },
        });
        return () =>
          h("div", {
            onInput: () => {
              hostInputs++;
            },
          });
      },
    }),
  );
  app.mount(container);
  try {
    const rendered = container.firstElementChild ?? panic("Expected mounted Vue DOM");
    expect(rendered.ownerDocument).toBe(document);
    expect(rendered instanceof HTMLElement).toBe(true);
    rendered.dispatchEvent(new InputEvent("input", { bubbles: true }));
    expect(hostInputs).toBe(1);
    const editor = holder.editor ?? panic("Expected mounted Vue editor");
    const bytes = await createDocx(createEmptyDocument({ initialText: "Start" }));
    await editor.loadBuffer(bytes);
    const detachedRead = editor.getDocument() ?? panic("Expected canonical document snapshot");
    detachedRead.package.document.content = [];
    expect(editor.getDocument()?.package.document.content).not.toHaveLength(0);
    const unsupportedWrite = editor.getDocument() ?? panic("Expected canonical document snapshot");
    unsupportedWrite.package.document.content = [];
    editor.setDocument(unsupportedWrite);
    expect(editor.getDocument()?.package.document.content).not.toHaveLength(0);
    expect(errors.at(0)?.message).toContain("Direct document model changes are unavailable");
    expect(editor.parseError.value).toBeNull();
    expect(editor.isReady.value).toBe(true);
    const view = editor.editorView.value ?? panic("Expected body editor view");
    // Exercise the hidden manager's refusal callback as well as model writes.
    const bold = view.state.schema.marks["bold"] ?? panic("Expected bold mark");
    view.dispatch(view.state.tr.addMark(1, 6, bold.create()));
    expect(errors.at(-1)?.message).toContain("unavailable");
    expect(editor.parseError.value).toBeNull();
    expect(editor.isReady.value).toBe(true);
    expect(view.state.doc.rangeHasMark(1, 6, bold)).toBe(false);
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 6)));
    view.dom.dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertText",
        data: " edited",
      }),
    );
    const canonical = editor.getDocument() ?? panic("Expected canonical document");
    expect(view.state.doc.textContent).toBe("Start edited");
    expect(editor.isDirty.value).toBe(true);
    await new Promise((resolve) => window.setTimeout(resolve, 300));
    expect(hostChanges).toBeGreaterThan(0);
    const saved = await editor.save();
    expect(errors).toHaveLength(2);
    expect(editor.parseError.value).toBeNull();
    if (!saved) panic("Expected saved canonical DOCX");
    const reopened = await parseDocx(await saved.arrayBuffer(), {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(reviewDifferences(canonical, reopened)).toEqual({ messages: [], omitted: 0 });
    expect(editor.isDirty.value).toBe(false);
  } finally {
    app.unmount();
    container.remove();
    hidden.remove();
    pages.remove();
  }
});

test("Vue canonical stories share history and save headers, first-page footer, notes and section properties", async () => {
  const container = document.createElement("div");
  const hidden = document.createElement("div");
  const pages = document.createElement("div");
  const notes = document.createElement("div");
  document.body.append(container, hidden, pages, notes);
  const errors: Error[] = [];
  const holder: { editor: ReturnType<typeof useDocxEditor> | null } = { editor: null };
  const app = createApp(
    defineComponent({
      setup() {
        holder.editor = useDocxEditor({
          hiddenContainer: shallowRef(hidden),
          pagesContainer: shallowRef(pages),
          noteEditorContainer: shallowRef(notes),
          experimentalSession: "canonical",
          featureFlags: () => ({ selectiveSave: true }),
          onError: (error) => errors.push(error),
        });
        return () => h("div");
      },
    }),
  );
  app.mount(container);
  try {
    const adapter = holder.editor ?? panic("Expected mounted Vue editor");
    const source =
      createEmptyHeaderFooter(createEmptyDocument({ initialText: "Body" }), "header", false) ??
      panic("Expected document with header");
    const rId =
      [...(source.package.headers?.keys() ?? [])].at(0) ?? panic("Expected header identity");
    await adapter.loadBuffer(await createDocx(source));
    const header = adapter.getHeaderFooterView(rId) ?? panic("Expected canonical header view");
    expect(
      header.state.doc.eq(
        adapter.editor.getCanonicalStoryProjection({ kind: "header", rId }) ??
          panic("Expected header projection"),
      ),
    ).toBe(true);
    header.dom.dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertText",
        data: "Header",
      }),
    );
    expect(errors).toEqual([]);
    expect(header.state.doc.textContent).toBe("Header");
    expect(adapter.isDirty.value).toBe(true);
    expect(
      adapter.editor.applyCanonicalStoryHistory({
        view: header,
        story: { kind: "header", rId },
        direction: "undo",
      }),
    ).toBe(true);
    expect(header.state.doc.textContent).toBe("");
    expect(
      adapter.editor.applyCanonicalStoryHistory({
        view: header,
        story: { kind: "header", rId },
        direction: "redo",
      }),
    ).toBe(true);
    expect(header.state.doc.textContent).toBe("Header");

    const beforeFooter = adapter.getDocument() ?? panic("Expected canonical document");
    const footerOp = createCanonicalHeaderFooterOperation({
      document: beforeFooter,
      position: "footer",
      referenceType: "first",
    });
    expect(adapter.editor.applyCanonicalOperations([footerOp])).toBe(true);
    const footer =
      adapter.getHeaderFooterView(footerOp.story.rId) ?? panic("Expected first-page footer view");
    footer.dom.dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertText",
        data: "First footer",
      }),
    );
    expect(footer.state.doc.textContent).toBe("First footer");

    const beforeNote = adapter.getDocument() ?? panic("Expected canonical document");
    const paragraph = beforeNote.package.document.content.at(0);
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
            content: withCanonicalParagraphIds([{ type: "paragraph", content: [] }], beforeNote),
          },
        },
      ]),
    ).toBe(true);
    adapter.openNoteStory({ kind: "footnote", noteId: 1 });
    const note = adapter.getActiveNoteView() ?? panic("Expected canonical footnote view");
    note.dom.dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertText",
        data: "Note",
      }),
    );
    expect(note.state.doc.textContent).toBe("Note");
    expect(
      adapter.editor.applyCanonicalStoryHistory({
        view: note,
        story: { kind: "footnote", id: 1 },
        direction: "undo",
      }),
    ).toBe(true);
    expect(note.state.doc.textContent).toBe("");
    expect(
      adapter.editor.applyCanonicalStoryHistory({
        view: note,
        story: { kind: "footnote", id: 1 },
        direction: "redo",
      }),
    ).toBe(true);
    expect(note.state.doc.textContent).toBe("Note");

    const controls = usePageSetupControls({
      editorView: adapter.editorView,
      getDocument: adapter.getDocument,
      readOnly: shallowRef(false),
      stateTick: shallowRef(0),
      reLayout: adapter.reLayout,
      onChange: () => panic("Canonical section writes must use the journal"),
      applySectionProperties: (properties) => {
        const current = adapter.getDocument() ?? panic("Expected canonical document");
        expect(
          adapter.editor.applyCanonicalOperations([
            createCanonicalSectionPropertiesOperation(current, properties),
          ]),
        ).toBe(true);
        return "applied";
      },
    });
    controls.handlePageSetupApply({ marginLeft: 720, footnotePr: { numStart: 2 } });
    const canonical = adapter.getDocument() ?? panic("Expected canonical document");
    expect(canonical.package.document.finalSectionProperties?.marginLeft).toBe(720);
    expect(canonical.package.document.finalSectionProperties?.titlePg).toBe(true);
    const saved = await adapter.save();
    if (!saved) panic("Expected canonical story save");
    const reopened = await parseDocx(await saved.arrayBuffer(), {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(reviewDifferences(canonical, reopened)).toEqual({ messages: [], omitted: 0 });
    expect(reopened.package.document.finalSectionProperties).toEqual(
      canonical.package.document.finalSectionProperties,
    );
    expect([...(reopened.package.headers?.keys() ?? [])].sort()).toEqual(
      [...(canonical.package.headers?.keys() ?? [])].sort(),
    );
    expect([...(reopened.package.footers?.keys() ?? [])].sort()).toEqual(
      [...(canonical.package.footers?.keys() ?? [])].sort(),
    );
    for (const [identity, part] of canonical.package.headers ?? []) {
      expect(canonicalReviewBlocks(reopened.package.headers?.get(identity)?.content ?? [])).toEqual(
        canonicalReviewBlocks(part.content),
      );
    }
    for (const [identity, part] of canonical.package.footers ?? []) {
      expect(canonicalReviewBlocks(reopened.package.footers?.get(identity)?.content ?? [])).toEqual(
        canonicalReviewBlocks(part.content),
      );
    }
    expect((reopened.package.footnotes ?? []).map(({ id }) => id)).toEqual(
      (canonical.package.footnotes ?? []).map(({ id }) => id),
    );
    for (const part of canonical.package.footnotes ?? []) {
      expect(
        canonicalReviewBlocks(
          reopened.package.footnotes?.find(({ id }) => id === part.id)?.content ?? [],
        ),
      ).toEqual(canonicalReviewBlocks(part.content));
    }
    expect(adapter.isDirty.value).toBe(false);
    expect(
      adapter.editor.applyCanonicalOperations(
        removeCanonicalHeaderFooterOperations({ document: canonical, position: "header", rId }),
      ),
    ).toBe(true);
    const withoutHeader = await adapter.save();
    if (!withoutHeader) panic("Expected saved header removal");
    const removed = await parseDocx(await withoutHeader.arrayBuffer(), {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(removed.package.headers?.has(rId) ?? false).toBe(false);
    expect(removed.package.document.finalSectionProperties?.headerReferences ?? []).toEqual([]);
    expect(errors).toEqual([]);
  } finally {
    app.unmount();
    container.remove();
    hidden.remove();
    pages.remove();
    notes.remove();
  }
});
