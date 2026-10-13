import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
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
