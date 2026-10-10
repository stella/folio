import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";
import type { Comment } from "@stll/folio-core/types/content";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "@stll/folio-core/document-operations";
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

// A notify-only controlled setter cannot publish requested metadata into the
// journal. Test both host decisions, including copied accepted prop values.
test.each(["accepted", "rejected"] as const)(
  "legacy comment-operation undo validates %s controlled metadata",
  async (decision) => {
    const container = document.createElement("div");
    document.body.append(container);
    const root = createRoot(container);
    const editor = createRef<DocxEditorRef>();
    const initialDocument = createEmptyDocument({ initialText: "Hello" });
    const comments: Comment[] = [];
    const notifications: Comment[][] = [];
    const controls = { comments, notifications };
    const onCommentsChange = (nextComments: Comment[]) => controls.notifications.push(nextComments);
    const render = () => {
      const props = {
        document: initialDocument,
        comments: controls.comments,
        onCommentsChange,
        showToolbar: false,
      };
      return (
        <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
          <DocxEditor ref={editor} {...props} />
        </IntlProvider>
      );
    };
    try {
      await act(async () => root.render(render()));
      await act(async () => editor.current?.ensureEditorView({ focus: false }));
      const ref = editor.current ?? panic("Editor ref did not mount");
      const view = ref.getEditorRef()?.getView() ?? panic("Body view did not mount");
      const before = view.state.doc;
      const snapshot = ref.createAIEditSnapshot() ?? panic("Snapshot was unavailable");
      const block = snapshot.blocks.at(0) ?? panic("Fixture has no block");
      const outcome: { result: ReturnType<DocxEditorRef["applyDocumentOperations"]> | null } = {
        result: null,
      };
      await act(async () => {
        outcome.result = ref.applyDocumentOperations({
          snapshot,
          batch: {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "direct",
            operations: [
              {
                id: "review",
                type: "commentOnBlock",
                blockId: block.id,
                comment: { text: "Review" },
              },
            ],
          },
        });
      });
      const applied = outcome.result ?? panic("Document operation did not return a result");
      expect(applied.status).toBe("committed");
      const handle =
        applied.undoHandle ?? panic("Document operation did not return an undo handle");
      const requested = controls.notifications.at(-1) ?? panic("Comment mutation was not notified");
      expect(requested).toHaveLength(1);
      // A request alone must not leak metadata into the public document output.
      expect(ref.getDocument()?.package.document.comments ?? []).toEqual([]);
      if (decision === "accepted") controls.comments = structuredClone(requested);
      await act(async () => root.render(render()));
      expect(ref.getDocument()?.package.document.comments ?? []).toEqual(
        decision === "accepted" ? requested : [],
      );
      await act(async () => {
        const undo = ref.undoDocumentOperations(handle);
        expect(undo.status).toBe(decision === "accepted" ? "undone" : "rejected");
        if (undo.status === "rejected") expect(undo.reason).toBe("documentChanged");
      });
      if (decision === "accepted") {
        expect(view.state.doc.eq(before)).toBe(true);
        const restored = controls.notifications.at(-1) ?? panic("Undo metadata was not notified");
        expect(restored).toEqual([]);
        expect(ref.getDocument()?.package.document.comments ?? []).toEqual(requested);
        controls.comments = structuredClone(restored);
        await act(async () => root.render(render()));
      } else {
        expect(view.state.doc.eq(before)).toBe(false);
      }
      expect(ref.getDocument()?.package.document.comments ?? []).toEqual([]);
    } finally {
      await act(async () => root.unmount());
      container.remove();
    }
  },
);
