import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { act, createRef, Profiler } from "react";
import { createRoot } from "react-dom/client";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { IntlProvider } from "use-intl";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const messages = getFolioMessages("en");
const createViewProbe = () => {
  const views: EditorView[] = [];
  return {
    views,
    onReady: (view: EditorView | null) => {
      if (view) views.push(view);
    },
  };
};
const createChangeProbe = () => {
  const changes: string[] = [];
  return {
    changes,
    onChange: (document: ReturnType<typeof createEmptyDocument>) =>
      changes.push(JSON.stringify(document.package.document.content)),
  };
};
const createRenderProbe = () => {
  const commits: number[] = [];
  return {
    commits,
    onRender: (_id: string, _phase: string, duration: number) => commits.push(duration),
  };
};

// Run this identical sequence on main and the change to compare committed
// render counts. Duration is diagnostic: machine load makes timing unsuitable
// as a correctness assertion. Initial mount and view initialization are excluded.
for (const experimentalSession of [undefined, "canonical"] as const) {
  test(`editor typing and selection render measurement (${experimentalSession ?? "default"})`, async () => {
    const container = document.createElement("div");
    document.body.append(container);
    const root = createRoot(container);
    const editor = createRef<DocxEditorRef>();
    const { views, onReady } = createViewProbe();
    const { commits, onRender } = createRenderProbe();
    const { changes, onChange } = createChangeProbe();
    const initialDocument = createEmptyDocument({ initialText: "Hello" });
    const typing: number[] = [];
    const selection: number[] = [];
    try {
      await act(async () => {
        root.render(
          <IntlProvider locale="en" timeZone="UTC" messages={messages}>
            <Profiler id="DocxEditor" onRender={onRender}>
              <DocxEditor
                ref={editor}
                document={initialDocument}
                {...(experimentalSession ? { experimentalSession } : {})}
                onEditorViewReady={onReady}
                onChange={onChange}
                showToolbar={false}
              />
            </Profiler>
          </IntlProvider>,
        );
      });
      await act(async () => {
        editor.current?.ensureEditorView({ focus: false });
      });
      const view = views.at(-1) ?? panic("The editor did not publish its body view");
      const startCommit = commits.length;
      const startChanges = changes.length;
      for (const character of " abcdef") {
        const before = commits.length;
        await act(async () => {
          view.dispatch(view.state.tr.insertText(character, view.state.doc.content.size - 1));
        });
        typing.push(commits.length - before);
      }
      expect(view.state.doc.textContent).toBe("Hello abcdef");
      // A selected range changes on every operation; this must exercise the
      // real selection callbacks rather than a no-op selection transaction.
      for (const position of [1, 2, 3, 4, 5, 6]) {
        const before = commits.length;
        await act(async () => {
          view.dispatch(
            view.state.tr.setSelection(
              TextSelection.create(view.state.doc, position, position + 1),
            ),
          );
        });
        selection.push(commits.length - before);
        expect(view.state.selection.from).toBe(position);
        expect(view.state.selection.to).toBe(position + 1);
      }
      expect(changes.length).toBeGreaterThan(startChanges);
      expect(typing.some((count) => count > 0)).toBe(true);
      console.info(
        "FOLIO_RENDER_MEASUREMENT",
        JSON.stringify({
          session: experimentalSession ?? "default",
          typing,
          selection,
          commits: commits.length - startCommit,
          durationMs: commits.slice(startCommit).reduce((sum, duration) => sum + duration, 0),
        }),
      );
    } finally {
      await act(async () => root.unmount());
      container.remove();
    }
  });
}
