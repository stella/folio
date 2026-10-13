import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import type {
  ComplexField,
  ParagraphContent,
  Run,
  SimpleField,
} from "@stll/folio-core/types/content";
import type { Document } from "@stll/folio-core/types/document";
import type { FindResult } from "@stll/folio-core/utils/findReplace";
import type { UseFindReplaceReturn } from "./useFindReplace";

const { createApp, defineComponent, h, ref, shallowRef } = await import("vue");
const { useFindReplace } = await import("./useFindReplace");

afterAll(() => GlobalRegistrator.unregister());

// The previous adapter searched only the body and selected the result in PM.
// Exercise host notifications and wraparound with a secondary story while
// asserting that search leaves both document content and editor selection intact.
test.each([true, false])("Vue search is view-only with readOnly=%s", (readOnly) => {
  const doc = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            content: [{ type: "run", content: [{ type: "text", text: "hit HIT" }] }],
          },
        ],
      },
      footnotes: [
        {
          type: "footnote",
          id: 1,
          content: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "hit" }] }],
            },
          ],
        },
      ],
    },
  } satisfies Document;
  const before = structuredClone(doc);
  const host = document.createElement("div");
  const hidden = document.createElement("div");
  document.body.append(host, hidden);
  const view = new EditorView(hidden, { state: EditorState.create({ doc: toProseDoc(doc) }) });
  const state = view.state;
  const notifications: (FindResult | null)[] = [];
  const holder: { search: UseFindReplaceReturn | null } = { search: null };
  const app = createApp(
    defineComponent({
      setup() {
        holder.search = useFindReplace({
          editorView: shallowRef(view),
          pagesRef: shallowRef(null),
          layout: shallowRef(null),
          getDocument: () => doc,
          displayMode: ref("no-markup"),
          readOnly: ref(readOnly),
          onResultChange: (result) => notifications.push(result),
        });
        return () => h("div");
      },
    }),
  );
  app.mount(host);
  try {
    const search = holder.search ?? panic("Search composable was not mounted");
    expect(search.find("hit")).toMatchObject({ totalCount: 3, currentIndex: 0 });
    expect(search.findNext()?.story).toBe("main");
    expect(search.findNext()?.story).toEqual({ kind: "footnote", id: 1 });
    expect(search.getFindResult()?.currentIndex).toBe(2);
    expect(search.findNext()?.story).toBe("main");
    expect(search.findPrevious()?.story).toEqual({ kind: "footnote", id: 1 });
    search.replaceText.value = "replacement";
    expect(search.replaceCurrent()).toBe(false);
    expect(search.find("hit", { matchCase: true })?.totalCount).toBe(2);
    expect(view.state).toBe(state);
    expect(doc).toEqual(before);
    search.clear();
    expect(search.getFindResult()).toBeNull();
    expect(notifications.at(-1)).toBeNull();
    expect(notifications.slice(0, 5).map((result) => result?.currentIndex)).toEqual([
      0, 1, 2, 0, 2,
    ]);
  } finally {
    app.unmount();
    view.destroy();
    host.remove();
    hidden.remove();
  }
});

const textRun = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });

const withReplacementEditor = (
  content: ParagraphContent[],
  check: (search: UseFindReplaceReturn, view: EditorView) => void,
) => {
  const base = {
    package: { document: { content: [{ type: "paragraph", content }] } },
  } satisfies Document;
  const host = document.createElement("div");
  const hidden = document.createElement("div");
  document.body.append(host, hidden);
  const view = new EditorView(hidden, { state: EditorState.create({ doc: toProseDoc(base) }) });
  const holder: { search: UseFindReplaceReturn | null } = { search: null };
  const app = createApp(
    defineComponent({
      setup() {
        holder.search = useFindReplace({
          editorView: shallowRef(view),
          pagesRef: shallowRef(null),
          layout: shallowRef(null),
          getDocument: () =>
            fromProseDoc(view.state.doc, base, { stylesheetSource: { type: "package" } }),
          displayMode: ref("no-markup"),
          readOnly: ref(false),
          onResultChange: () => {},
        });
        return () => h("div");
      },
    }),
  );
  app.mount(host);
  try {
    check(holder.search ?? panic("Search composable was not mounted"), view);
  } finally {
    app.unmount();
    view.destroy();
    host.remove();
    hidden.remove();
  }
};

const cachedFields = {
  simple: {
    type: "simpleField",
    fieldType: "REF",
    instruction: "REF target",
    content: [textRun("42")],
  },
  complex: {
    type: "complexField",
    fieldType: "REF",
    instruction: "REF target",
    fieldCode: [],
    fieldResult: [textRun("42")],
  },
} satisfies Record<string, SimpleField | ComplexField>;

// Search fixtures previously checked projection and offsets without executing
// adapter transactions. Exercise both field forms and mixed eligible result
// sets against real PM state so a field atom can never be deleted by replace.
test.each(Object.entries(cachedFields))(
  "Vue refuses replacing a %s field substring",
  (_name, field) => {
    withReplacementEditor([field], (search, view) => {
      const before = view.state;
      expect(search.find("4")?.totalCount).toBe(1);
      expect(search.currentMatch.value?.replacement).toEqual({
        status: "blocked",
        reason: "field-result",
      });
      expect(search.canReplaceCurrent.value).toBe(false);
      search.replaceText.value = "9";
      expect(search.replaceCurrent()).toBe(false);
      expect(search.replaceAll()).toEqual({ replacedCount: 0, skippedCount: 1 });
      expect(view.state).toBe(before);
    });
  },
);

test.each(["fields-only", "mixed"] as const)(
  "Vue replace all reports skipped fields in a %s result set",
  (kind) => {
    const content: ParagraphContent[] = [cachedFields.simple, cachedFields.complex];
    if (kind === "mixed") content.push(textRun(" 4"));
    withReplacementEditor(content, (search, view) => {
      const fieldsBefore: unknown[] = [];
      view.state.doc.descendants((node) => {
        if (node.type.name === "field") fieldsBefore.push(node.toJSON());
      });
      expect(fieldsBefore).toHaveLength(2);
      const before = view.state;
      expect(search.find("4")?.totalCount).toBe(kind === "mixed" ? 3 : 2);
      search.replaceText.value = "9";
      expect(search.replaceAll()).toEqual({
        replacedCount: kind === "mixed" ? 1 : 0,
        skippedCount: 2,
      });
      const fieldsAfter: unknown[] = [];
      view.state.doc.descendants((node) => {
        if (node.type.name === "field") fieldsAfter.push(node.toJSON());
      });
      expect(fieldsAfter).toEqual(fieldsBefore);
      if (kind === "fields-only") expect(view.state).toBe(before);
      else expect(view.state.doc.textContent).toBe("4242 9");
    });
  },
);

test("Vue refuses a replacement crossing hidden deleted content", () => {
  withReplacementEditor(
    [
      textRun("A"),
      { type: "deletion", info: { id: 1, author: "Author" }, content: [textRun("X")] },
      textRun("B"),
    ],
    (search, view) => {
      const before = view.state;
      expect(search.find("AB")?.totalCount).toBe(1);
      expect(search.currentMatch.value?.replacement).toEqual({
        status: "blocked",
        reason: "hidden-content",
      });
      expect(search.canReplaceCurrent.value).toBe(false);
      search.replaceText.value = "C";
      expect(search.replaceCurrent()).toBe(false);
      expect(search.replaceAll()).toEqual({ replacedCount: 0, skippedCount: 1 });
      expect(view.state).toBe(before);
      expect(view.state.doc.textContent).toBe("AXB");
    },
  );
});

test("Vue replaces an ordinary editable match", () => {
  withReplacementEditor([textRun("before hit after")], (search, view) => {
    expect(search.find("hit")?.totalCount).toBe(1);
    expect(search.currentMatch.value?.replacement).toEqual({ status: "allowed" });
    expect(search.canReplaceCurrent.value).toBe(true);
    search.replaceText.value = "changed";
    expect(search.replaceCurrent()).toBe(true);
    expect(view.state.doc.textContent).toBe("before changed after");
    expect(search.getFindResult()?.totalCount).toBe(0);
  });
});

// A debounced refresh can be pending when a host changes the current PM doc.
// Re-check the result metadata before any transaction crosses that boundary.
test.each(["current", "all"] as const)(
  "Vue refreshes eligibility before %s replacement after an edit",
  (command) => {
    withReplacementEditor([textRun("42")], (search, view) => {
      expect(search.find("4")?.totalCount).toBe(1);
      expect(search.currentMatch.value?.replacement).toEqual({ status: "allowed" });
      const fieldDoc = toProseDoc({
        package: { document: { content: [{ type: "paragraph", content: [cachedFields.simple] }] } },
      });
      view.dispatch(view.state.tr.replaceWith(0, view.state.doc.content.size, fieldDoc.content));
      const before = view.state;
      search.replaceText.value = "9";
      if (command === "current") expect(search.replaceCurrent()).toBe(false);
      else expect(search.replaceAll()).toEqual({ replacedCount: 0, skippedCount: 1 });
      expect(search.currentMatch.value?.replacement).toEqual({
        status: "blocked",
        reason: "field-result",
      });
      expect(view.state).toBe(before);
    });
  },
);
