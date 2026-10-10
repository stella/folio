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
import { createDocx } from "@stll/folio-core/docx/rezip";
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

// Identical default-session sequence measured on main and this change in CI:
// 2 commits per keystroke/selection, then 4 commits after projection settles.
const DEFAULT_RENDER_BUDGET = { perOperation: 2, settled: 4, total: 30 };

// Run this identical sequence on main and the change to compare committed
// render counts. Duration is diagnostic: machine load makes timing unsuitable
// as a correctness assertion. Initial mount and view initialization are excluded.
for (const experimentalSession of [undefined, "canonical"] as const) {
  test(`editor typing and selection render measurement (${experimentalSession ?? "default"})`, async () => {
    const container = document.createElement("div");
    document.body.append(container);
    const root = createRoot(container);
    const editor = createRef<DocxEditorRef>();
    const { onReady } = createViewProbe();
    const { commits, onRender } = createRenderProbe();
    const { changes, onChange } = createChangeProbe();
    const initialDocument = createEmptyDocument({ initialText: "Hello" });
    // Load package bytes explicitly, as in the canonical-save harness, so
    // the canonical source and body projection are ready before measuring.
    const documentBuffer = experimentalSession ? await createDocx(initialDocument) : undefined;
    const documentProps = documentBuffer
      ? { documentBuffer, experimentalSession }
      : { document: initialDocument };
    const typing: number[] = [];
    const selection: number[] = [];
    try {
      await act(async () => {
        root.render(
          <IntlProvider locale="en" timeZone="UTC" messages={messages}>
            <Profiler id="DocxEditor" onRender={onRender}>
              <DocxEditor
                ref={editor}
                {...documentProps}
                onEditorViewReady={onReady}
                onChange={onChange}
                showToolbar={false}
              />
            </Profiler>
          </IntlProvider>,
        );
      });
      if (documentBuffer) {
        await act(async () => editor.current?.loadDocumentBuffer(documentBuffer));
      }
      await act(async () => {
        editor.current?.ensureEditorView({ focus: false });
      });
      // Let mount-time font/layout and history notifications settle before measuring.
      await act(async () => {
        await new Promise((resolve) => setTimeout(resolve, 450));
      });
      const api = editor.current?.getEditor() ?? panic("The editor API did not mount");
      const view = api.getView() ?? panic("The editor did not create its body view");
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
      const immediateCommits = commits.length;
      // Document projection is deliberately debounced off the keypress path.
      await act(async () => {
        await new Promise((resolve) => setTimeout(resolve, 800));
      });
      expect(changes.length).toBeGreaterThan(startChanges);
      expect(typing.some((count) => count > 0)).toBe(true);
      if (!experimentalSession) {
        for (const count of [...typing, ...selection]) {
          expect(count).toBeLessThanOrEqual(DEFAULT_RENDER_BUDGET.perOperation);
        }
        expect(commits.length - immediateCommits).toBeLessThanOrEqual(
          DEFAULT_RENDER_BUDGET.settled,
        );
        expect(commits.length - startCommit).toBeLessThanOrEqual(DEFAULT_RENDER_BUDGET.total);
      }
      console.info(
        "FOLIO_RENDER_MEASUREMENT",
        JSON.stringify({
          session: experimentalSession ?? "default",
          typing,
          selection,
          commits: commits.length - startCommit,
          settledCommits: commits.length - immediateCommits,
          durationMs: commits.slice(startCommit).reduce((sum, duration) => sum + duration, 0),
        }),
      );
    } finally {
      await act(async () => root.unmount());
      container.remove();
    }
  });
}
