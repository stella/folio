import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { TextSelection } from "prosemirror-state";

const { createApp, defineComponent, h, shallowRef } = await import("vue");

import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { CanonicalDocxInputError } from "@stll/folio-core/docx/canonicalSessionInput";
import { reviewDifferences } from "../../../../test/reviewDifferences";

const { useDocxEditor } = await import("./useDocxEditor");

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
