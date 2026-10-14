import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { DISPLAY_MODES } from "@stll/folio-core/managers/EditorModeManager";
import type { DisplayMode } from "@stll/folio-core/managers/EditorModeManager";
import { schema } from "@stll/folio-core/prosemirror/schema";
import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import { repackDocx } from "@stll/folio-core/docx/rezip";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import { buildRawPackage, SECTION, WORD_STYLES } from "@stll/folio-core/__tests__/documentShapes";
import type { EditorReadOnlyReason } from "@stll/folio-core/managers/EditorModeManager";

const { effectScope, nextTick, ref, shallowRef } = await import("vue");
const { useFindReplace } = await import("./useFindReplace");

afterAll(() => GlobalRegistrator.unregister());

test("initial projected searches defer until the lazy body editor mounts", () => {
  const revisionAttrs = { revisionId: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" };
  const insertion = (schema.marks["insertion"] ?? panic("Missing insertion mark")).create(
    revisionAttrs,
  );
  const deletion = (schema.marks["deletion"] ?? panic("Missing deletion mark")).create(
    revisionAttrs,
  );
  const state = EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("inserted", [insertion]),
        schema.text(" deleted", [deletion]),
        schema.text(" kept"),
      ]),
    ]),
  });
  const hiddenQueries = {
    "all-markup": null,
    "simple-markup": "deleted",
    "no-markup": "deleted",
    original: "inserted",
  } as const satisfies Record<DisplayMode, string | null>;
  for (const mode of DISPLAY_MODES) {
    const hiddenQuery = hiddenQueries[mode];
    if (!hiddenQuery) continue;
    const editorView = shallowRef<EditorView | null>(null);
    const scope = effectScope();
    const find =
      scope.run(() => useFindReplace({ editorView, markupView: ref(mode) })) ??
      panic("Expected Vue find composable");
    let view: EditorView | null = null;
    try {
      find.searchText.value = hiddenQuery;
      expect(find.performSearch()).toEqual([]);
      expect(find.findNext()).toBeNull();
      view = new EditorView(document.createElement("div"), { state });
      editorView.value = view;
      expect(find.performSearch()).toEqual([]);
      find.searchText.value = "kept";
      expect(find.performSearch()).toHaveLength(1);
    } finally {
      scope.stop();
      view?.destroy();
    }
  }
});

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
    expect(find.replaceCurrent()).toBe(false);
    expect(find.replaceAll()).toBe(0);
    markupView.value = "all-markup";
    await nextTick();
    expect(find.replaceCurrent()).toBe(true);
    expect(view.state.doc.textContent).toBe("replacement kept");
  } finally {
    scope.stop();
    view.destroy();
  }
});

test("projected replacement refuses a match spanning hidden revisions without changing saved content", async () => {
  const input = await buildRawPackage({
    styles: WORD_STYLES,
    body:
      '<w:p><w:r><w:t>left</w:t></w:r><w:del w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve"> hidden </w:delText></w:r></w:del><w:r><w:t>right</w:t></w:r></w:p>' +
      SECTION(),
  });
  const source = await parseDocx(input, { preloadFonts: false, detectVariables: false });
  const doc = toProseDoc(source);
  const view = new EditorView(document.createElement("div"), {
    state: EditorState.create({ doc }),
    editable: () => true,
  });
  const markupView = ref<DisplayMode>("no-markup");
  const reasons: EditorReadOnlyReason[] = [];
  const scope = effectScope();
  const find =
    scope.run(() =>
      useFindReplace({
        editorView: shallowRef(view),
        markupView,
        readOnly: ref(false),
        onReadonlyEditAttempt: (reason) => reasons.push(reason),
      }),
    ) ?? panic("Expected Vue find composable");
  const originalJson = JSON.stringify(doc.toJSON());
  const serializeLiveDocument = async () =>
    new Uint8Array(
      await repackDocx(
        fromProseDoc(view.state.doc, source, { stylesheetSource: { type: "package" } }),
        { updateModifiedDate: false },
      ),
    );
  const savedBefore = await serializeLiveDocument();
  try {
    find.searchText.value = "leftright";
    expect(find.performSearch()).toMatchObject([{ from: 1, to: 18, text: "leftright" }]);
    find.replaceText.value = "replacement";
    expect(find.replaceCurrent()).toBe(false);
    expect(find.replaceAll()).toBe(0);
    expect(reasons).toEqual(["review-display-mode", "review-display-mode"]);
    expect(JSON.stringify(view.state.doc.toJSON())).toBe(originalJson);
    const savedAfter = await serializeLiveDocument();
    expect(savedAfter).toEqual(savedBefore);
    markupView.value = "all-markup";
    await nextTick();
    find.searchText.value = "left";
    expect(find.performSearch()).toHaveLength(1);
    expect(find.replaceCurrent()).toBe(true);
    expect(view.state.doc.textContent).toBe("replacement hidden right");
  } finally {
    scope.stop();
    view.destroy();
  }
});
