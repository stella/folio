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
import { TextSelection } from "prosemirror-state";

const { createApp, defineComponent, h, shallowRef } = await import("vue");

import { createDocx } from "@stll/folio-core/docx/rezip";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { setDocumentWatermark } from "@stll/folio-core/watermark";

const { documentChrome } = await import("./documentChrome");
const { useDocxEditor } = await import("./useDocxEditor");

afterAll(() => GlobalRegistrator.unregister());

// Longer than the change notification's debounce.
const SETTLE_MS = 400;
const settle = () => new Promise((resolve) => setTimeout(resolve, SETTLE_MS));

const typeInto = (dom: HTMLElement, data: string) =>
  dom.dispatchEvent(
    new InputEvent("beforeinput", {
      bubbles: true,
      cancelable: true,
      inputType: "insertText",
      data,
    }),
  );

/** A small deterministic generator, so a failing seed replays. */
const generator = (seed: number) => {
  let state = seed;
  return (bound: number) => {
    state = (state * 1_103_515_245 + 12_345) % 2 ** 31;
    return Math.floor((state / 2 ** 31) * bound);
  };
};

const source = async () =>
  createDocx(
    setDocumentWatermark(
      createEmptyDocument({ initialText: "First clause of the agreement." }),
      { kind: "text", text: "DRAFT" },
    ),
  );

const SEEDS = [1, 7, 42, 2026];

for (const engine of ["legacy", "canonical"] as const) {
  test.each(SEEDS)(
    `${engine}: once edits settle, the model's toolbar values equal those of a built document (seed %i)`,
    async (seed) => {
      const container = document.createElement("div");
      const hidden = document.createElement("div");
      const pages = document.createElement("div");
      document.body.append(container, hidden, pages);
      const holder: { editor: ReturnType<typeof useDocxEditor> | null } = {
        editor: null,
      };
      const app = createApp(
        defineComponent({
          setup() {
            holder.editor = useDocxEditor({
              hiddenContainer: shallowRef(hidden),
              pagesContainer: shallowRef(pages),
              experimentalSession: engine,
              onError: (error) => panic(error.message),
            });
            return () => h("div");
          },
        }),
      );
      app.mount(container);
      try {
        const editor = holder.editor ?? panic("Expected mounted Vue editor");
        await editor.loadBuffer(await source());
        editor.focus();
        const built = () =>
          documentChrome(editor.getDocument() ?? panic("Expected a document"));
        const model = () => documentChrome(editor.modelDocument.value);
        expect(model().watermark).toMatchObject({
          kind: "text",
          text: "DRAFT",
        });
        expect(model()).toEqual(built());

        const textBefore = editor.editorView.value?.state.doc.textContent;
        const loadedBuffer = editor.sourceBuffer.value;
        expect(loadedBuffer).toBeInstanceOf(ArrayBuffer);
        const random = generator(seed);
        const edits = 3 + random(6);
        for (let edit = 0; edit < edits; edit += 1) {
          const view = editor.editorView.value ?? panic("Expected the view");
          const size = view.state.doc.content.size;
          const at = 1 + random(Math.max(1, size - 2));
          view.dispatch(
            view.state.tr.setSelection(
              TextSelection.near(view.state.doc.resolve(at)),
            ),
          );
          if (engine === "canonical" && random(3) === 0) {
            // A document-wide value changes through the session too.
            editor.editor.applyCanonicalSectionProperties({
              marginLeft: 720 + random(4) * 180,
            });
          } else {
            typeInto(view.dom, "abc".slice(0, 1 + random(3)));
          }
          // A canonical model is the session's own document, so it holds
          // between keys too; a written-back model holds once edits settle.
          if (engine === "canonical") expect(model()).toEqual(built());
        }
        await settle();
        // The edits reached the document, so there was something to settle.
        expect(editor.editorView.value?.state.doc.textContent).not.toBe(
          textBefore,
        );
        expect(model()).toEqual(built());
        // Embedded fonts load once per document: edits keep its source bytes.
        expect(editor.sourceBuffer.value).toBe(loadedBuffer);
      } finally {
        app.unmount();
        container.remove();
        hidden.remove();
        pages.remove();
      }
    },
  );
}
