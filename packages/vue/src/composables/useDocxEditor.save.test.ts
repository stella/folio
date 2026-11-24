import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { panic } from "better-result";
import { afterAll, expect, test } from "bun:test";

const { createApp, defineComponent, h, shallowRef } = await import("vue");

import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

const { useDocxEditor } = await import("./useDocxEditor");

// Earlier save coverage had no edit between serialization and completion. Vary
// that sequence against the real composable and reopened DOCX output.
afterAll(() => GlobalRegistrator.unregister());

test.each([undefined, "canonical"] as const)(
  "%s autocomplete callbacks bind before readiness and report teardown",
  async (session) => {
    const container = document.createElement("div");
    const hidden = document.createElement("div");
    const pages = document.createElement("div");
    document.body.append(container, hidden, pages);
    const calls: string[] = [];
    const holder: { adapter: ReturnType<typeof useDocxEditor> | null } = {
      adapter: null,
    };
    let label = "first";
    const app = createApp(
      defineComponent({
        setup() {
          holder.adapter = useDocxEditor({
            hiddenContainer: shallowRef(hidden),
            pagesContainer: shallowRef(pages),
            ...(session === undefined ? {} : { experimentalSession: session }),
            onAutocompleteChange: (state) =>
              calls.push(
                `${label}:${state.status === "idle" ? "idle" : state.text}`,
              ),
            onError: (error) => panic(error.message),
          });
          return () => h("div");
        },
      }),
    );
    app.mount(container);
    try {
      const adapter = holder.adapter ?? panic("Expected mounted composable");
      expect(adapter.editorView.value).toBeNull();
      const api = adapter.editor.autocomplete;
      for (label of ["first", "replacement"]) {
        adapter.loadDocument(createEmptyDocument({ initialText: label }));
        expect(adapter.editor.autocomplete).toBe(api);
        expect(adapter.editorView.value).not.toBeNull();
        expect(api.start({ requestId: label })).toBe(label);
        api.append(label, " next");
        expect(calls.at(-1)).toBe(`${label}: next`);
      }
      expect(calls).toContain("replacement:idle");
      api.clear();
      expect(calls.at(-1)).toBe("replacement:idle");
      api.start({ requestId: "teardown" });
      api.append("teardown", " final");
      expect(calls.at(-1)).toBe("replacement: final");
    } finally {
      app.unmount();
      container.remove();
      hidden.remove();
      pages.remove();
    }
    expect(calls.at(-1)).toBe("replacement:idle");
  },
);
