import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";

const { createApp, defineComponent, h, shallowRef } = await import("vue");

import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { dispatchEditorTextInput } from "@stll/folio-core/prosemirror/textInput";
import { reviewDifferences } from "../../../../test/reviewDifferences";

const { useDocxEditor } = await import("./useDocxEditor");

// The save oracle exercises the composable's real hidden manager and serialization,
// rather than rebuilding the expected document from its PM projection.
afterAll(() => GlobalRegistrator.unregister());

test("canonical edits save and reopen the canonical text and paragraph identity", async () => {
  const container = document.createElement("div");
  const hidden = document.createElement("div");
  const pages = document.createElement("div");
  document.body.append(container, hidden, pages);
  const errors: Error[] = [];
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
        });
        return () => h("div");
      },
    }),
  );
  app.mount(container);
  try {
    const editor = holder.editor ?? panic("Expected mounted Vue editor");
    const bytes = await createDocx(createEmptyDocument({ initialText: "Start" }));
    await editor.loadBuffer(bytes);
    const view = editor.editorView.value ?? panic("Expected body editor view");
    dispatchEditorTextInput(view, { from: 6, to: 6, text: " edited" });
    const canonical = editor.getDocument() ?? panic("Expected canonical document");
    expect(view.state.doc.textContent).toBe("Start edited");
    expect(editor.isDirty.value).toBe(true);
    const saved = await editor.save();
    expect(errors).toEqual([]);
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
