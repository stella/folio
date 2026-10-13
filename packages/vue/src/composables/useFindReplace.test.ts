import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import type { DisplayMode } from "@stll/folio-core/managers/EditorModeManager";
import { schema } from "@stll/folio-core/prosemirror/schema";

const { effectScope, nextTick, ref, shallowRef } = await import("vue");
const { useFindReplace } = await import("./useFindReplace");

afterAll(() => GlobalRegistrator.unregister());

test("Vue searches refresh with the markup view and replacement respects read-only state", async () => {
  const deletion = (schema.marks["deletion"] ?? panic("Missing deletion mark")).create({
    revisionId: 1,
    author: "Reviewer",
    date: "2026-01-01T00:00:00Z",
  });
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [schema.text("needle", [deletion]), schema.text(" kept")]),
  ]);
  const view = new EditorView(document.createElement("div"), {
    state: EditorState.create({ doc }),
  });
  const markupView = ref<DisplayMode>("all-markup");
  const readOnly = ref(true);
  const scope = effectScope();
  const find =
    scope.run(() => useFindReplace({ editorView: shallowRef(view), markupView, readOnly })) ??
    panic("Expected Vue find composable");
  try {
    find.searchText.value = "needle";
    expect(find.performSearch()).toHaveLength(1);
    markupView.value = "no-markup";
    await nextTick();
    expect(find.matches.value).toEqual([]);
    markupView.value = "original";
    await nextTick();
    expect(find.matches.value).toHaveLength(1);
    expect(find.currentMatch.value).toMatchObject({ from: 1, to: 7 });
    find.replaceText.value = "replacement";
    expect(find.replaceCurrent()).toBe(false);
    expect(find.replaceAll()).toBe(0);
    expect(view.state.doc).toBe(doc);

    readOnly.value = false;
    view.setProps({ editable: () => false });
    expect(find.replaceCurrent()).toBe(false);
    expect(find.replaceAll()).toBe(0);
    expect(view.state.doc).toBe(doc);

    view.setProps({ editable: () => true });
    expect(find.replaceCurrent()).toBe(true);
    expect(view.state.doc.textContent).toBe("replacement kept");
  } finally {
    scope.stop();
    view.destroy();
  }
});
