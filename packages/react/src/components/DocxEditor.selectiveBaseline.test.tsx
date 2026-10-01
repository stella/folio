import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { ensureParaIds } from "@stll/folio-core/docx/ensureParaIds";
import type { TripwireResult } from "@stll/folio-core/docx/selectiveSaveTripwire";
import { fromMarkdown } from "@stll/folio-core/markdown";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";

import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const featureFlags = { selectiveSaveTripwire: true };

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

// Earlier save tests exercised one save against the loaded package. This sequence
// changes a different paragraph after every save and varies mid-save edits,
// exercising baseline ownership and the serialization/completion sequence.
test.each(["settled", "overlapping"])(
  "%s selective saves preserve previous and concurrent edits",
  async (timing) => {
    const container = document.createElement("div");
    document.body.append(container);
    const root = createRoot(container);
    const editor = createRef<DocxEditorRef>();
    const bytes = (await ensureParaIds(await createDocx(fromMarkdown("Alpha\n\nBeta")))).docx;
    const errors: Error[] = [];
    const tripwires: TripwireResult[] = [];
    // oxlint-disable-next-line react-perf/jsx-no-new-function-as-prop -- This test renders the editor once.
    const onSelectiveSaveTripwire = (result: TripwireResult) => tripwires.push(result);
    // oxlint-disable-next-line react-perf/jsx-no-new-function-as-prop -- This test renders the editor once.
    const onError = (error: Error) => errors.push(error);
    try {
      await act(async () => {
        root.render(
          <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
            <DocxEditor
              ref={editor}
              documentBuffer={bytes}
              onError={onError}
              featureFlags={featureFlags}
              onSelectiveSaveTripwire={onSelectiveSaveTripwire}
              showToolbar={false}
            />
          </IntlProvider>,
        );
      });
      await act(async () => editor.current?.loadDocumentBuffer(bytes));
      await act(async () => editor.current?.ensureEditorView({ focus: false }));
      const view = editor.current?.getEditor()?.getView() ?? panic("Expected body editor view");
      const expected = ["Alpha first", "Beta second"];
      for (const [index, suffix] of [" first", " second"].entries()) {
        await act(async () => {
          const paragraph = view.state.doc.child(index);
          const start = index === 0 ? 0 : view.state.doc.child(0).nodeSize;
          view.dispatch(view.state.tr.insertText(suffix, start + paragraph.nodeSize - 1));
          // Publish the edited Document before save, so the loader's effect sees
          // the round-tripped source package and can overwrite a stale baseline.
          editor.current?.getDocument();
        });
        let saved: ArrayBuffer | null | undefined;
        await act(async () => {
          const saving = editor.current?.save();
          if (timing === "overlapping") {
            const paragraph = view.state.doc.child(index);
            const start = index === 0 ? 0 : view.state.doc.child(0).nodeSize;
            view.dispatch(view.state.tr.insertText(" concurrent", start + paragraph.nodeSize - 1));
          }
          saved = await saving;
        });
        if (!saved) panic("Expected saved DOCX");
        expect(tripwires.at(-1)?.kind).not.toBe("selective-skipped");
        const reopened = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
        const projected = toProseDoc(reopened);
        for (let savedIndex = 0; savedIndex <= index; savedIndex++) {
          expect(projected.child(savedIndex).textContent).toBe(expected.at(savedIndex));
        }
        if (timing === "overlapping") {
          expect(editor.current?.hasPendingChanges()).toBe(true);
          expected[index] += " concurrent";
          await act(async () => {
            const completed = await editor.current?.save();
            if (!completed) panic("Expected completed save");
            const latest = await parseDocx(completed, {
              preloadFonts: false,
              detectVariables: false,
            });
            expect(toProseDoc(latest).child(index).textContent).toBe(expected.at(index));
          });
          expect(editor.current?.hasPendingChanges()).toBe(false);
        }
      }
      expect(errors).toEqual([]);
    } finally {
      await act(async () => root.unmount());
      container.remove();
    }
  },
);
