import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";

const { createApp, defineComponent, h, shallowRef } = await import("vue");

import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { dispatchEditorTextInput } from "@stll/folio-core/prosemirror/textInput";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";

const { useDocxEditor } = await import("./useDocxEditor");

// Earlier save coverage had no edit between serialization and completion. Vary
// that sequence against the real composable and reopened DOCX output.
afterAll(() => GlobalRegistrator.unregister());

test.each(["settled", "overlapping"])(
  "%s Vue saves preserve edits made during serialization",
  async (timing) => {
    const container = document.createElement("div");
    const hidden = document.createElement("div");
    const pages = document.createElement("div");
    document.body.append(container, hidden, pages);
    const errors: Error[] = [];
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
            onError: (error) => errors.push(error),
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
      const view = editor.editorView.value ?? panic("Expected body editor view");
      dispatchEditorTextInput(view, { from: 6, to: 6, text: " edited" });
      expect(view.state.doc.textContent).toBe("Start edited");
      expect(editor.isDirty.value).toBe(true);
      const saving = editor.save();
      if (timing === "overlapping") {
        view.dispatch(view.state.tr.insertText(" concurrent", view.state.doc.content.size - 1));
      }
      const saved = await saving;
      expect(errors).toEqual([]);
      if (!saved) panic("Expected saved DOCX");
      const reopened = await parseDocx(await saved.arrayBuffer(), {
        preloadFonts: false,
        detectVariables: false,
      });
      expect(toProseDoc(reopened).textContent).toBe("Start edited");
      expect(editor.isDirty.value).toBe(timing === "overlapping");
      if (timing === "overlapping") {
        expect(editor.getDocument()).not.toBeNull();
        expect(view.state.doc.textContent).toBe("Start edited concurrent");
        const completed = await editor.save();
        if (!completed) panic("Expected completed save");
        const latest = await parseDocx(await completed.arrayBuffer(), {
          preloadFonts: false,
          detectVariables: false,
        });
        expect(toProseDoc(latest).textContent).toBe("Start edited concurrent");
        expect(editor.isDirty.value).toBe(false);
      }
    } finally {
      app.unmount();
      container.remove();
      hidden.remove();
      pages.remove();
    }
  },
);
