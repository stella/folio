import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, setSystemTime, test } from "bun:test";
import { panic } from "better-result";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import { EditorState, TextSelection } from "prosemirror-state";
import { IntlProvider } from "use-intl";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { createCanonicalSession } from "@stll/folio-core/controller/canonicalSession";
import { schema } from "@stll/folio-core/prosemirror/schema";
import { shapeArrayBuffer } from "../../../core/src/__tests__/documentShapes";
import { assertExactModel } from "../../../../test/exactModel";
import type { Comment } from "@stll/folio-core/types/content";
import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const createErrorObserver = () => {
  const errors: Error[] = [];
  return { errors, onError: (error: Error) => errors.push(error) };
};

const SHAPES = ["header-footer", "mixed-lists", "single-decimal-list", "image"] as const;
const MODES = ["editing", "suggesting"] as const;
const COMPLETIONS = ["cancel", "commit"] as const;

for (const shape of SHAPES) {
  for (const mode of MODES) {
    for (const completion of COMPLETIONS) {
      test(`canonical comments render during ${shape} ${mode} composition ${completion}`, async () => {
        const container = document.createElement("div");
        document.body.append(container);
        const root = createRoot(container);
        const editor = createRef<DocxEditorRef>();
        const bytes = await shapeArrayBuffer(shape);
        const messages = getFolioMessages("en");
        const observer = createErrorObserver();
        const controls: { comments: Comment[]; showToolbar: boolean } = {
          comments: [],
          showToolbar: false,
        };
        const renderEditor = () => (
          <IntlProvider locale="en" timeZone="UTC" messages={messages}>
            <DocxEditor
              ref={editor}
              documentBuffer={bytes}
              experimentalSession="canonical"
              comments={controls.comments}
              onError={observer.onError}
              showToolbar={controls.showToolbar}
            />
          </IntlProvider>
        );
        try {
          await act(async () => root.render(renderEditor()));
          await act(async () => editor.current?.loadDocumentBuffer(bytes));
          await act(async () => editor.current?.ensureEditorView({ focus: false }));
          const api = editor.current?.getEditor() ?? panic("Expected canonical editor");
          const view = api.getView() ?? panic("Expected mounted canonical view");
          const sessionMode =
            mode === "editing"
              ? { type: "editing" as const }
              : { type: "suggesting" as const, author: "Composition author" };
          expect(api.setCanonicalMode(sessionMode)).toBe(true);
          await act(async () => {
            view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 4)));
          });
          const baseline = api.getCanonicalDocument() ?? panic("Expected committed model");
          const baselineProjection = view.state.doc;
          setSystemTime(new Date("2026-10-05T12:00:00Z"));
          const expected = createCanonicalSession(baseline).unwrap();
          expected.setMode(sessionMode);
          const expectedState = EditorState.create({ schema, doc: expected.projection.doc });
          if (completion === "commit") {
            expected
              .prepareReplace(expectedState, {
                from: 1,
                to: 4,
                text: "alpha",
                semantic: "composition",
              })
              .unwrap()
              .publish()
              .unwrap();
          }
          await act(async () => {
            view.dom.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true }));
            view.dispatch(view.state.tr.insertText("alpha", 1, 4).setMeta("composition", 1));
          });
          expect(api.getCanonicalDocument).toThrow(
            "Composition must finish before taking a snapshot.",
          );
          // Selection notifications and a controlled-prop render must read only
          // published comments, while public snapshots retain their refusal.
          await act(async () => {
            view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 2)));
          });
          controls.comments = [...controls.comments];
          controls.showToolbar = true;
          await act(async () => root.render(renderEditor()));
          expect(editor.current?.getEditor()?.getView()).toBe(view);
          expect(view.isDestroyed).toBe(false);
          expect(api.getCanonicalDocument).toThrow(
            "Composition must finish before taking a snapshot.",
          );
          await act(async () => {
            if (completion === "cancel") {
              // Native cancellation restores the selected source before its
              // final flush; PM swallows synthetic Escape while composing.
              view.dispatch(
                view.state.tr
                  .insertText(baselineProjection.textBetween(1, 4), 1, 6)
                  .setMeta("composition", 1),
              );
              expect(view.state.doc.eq(baselineProjection)).toBe(true);
            }
            view.dom.dispatchEvent(new CompositionEvent("compositionend", { bubbles: true }));
            await new Promise<void>((resolve) => setTimeout(resolve, 40));
          });
          const committed = api.getCanonicalDocument() ?? panic("Expected completed snapshot");
          assertExactModel(committed, expected.document);
          expect(editor.current?.getEditor()?.getView()).toBe(view);
          expect(observer.errors).toEqual([]);
        } finally {
          setSystemTime();
          await act(async () => root.unmount());
          container.remove();
        }
      });
    }
  }
}
