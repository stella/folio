import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { panic } from "better-result";
import { afterAll, expect, mock, test } from "bun:test";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import { schema } from "@stll/folio-core/prosemirror/schema";
import type { Document } from "@stll/folio-core/types/document";

const { createApp, defineComponent, h, ref, shallowRef } = await import("vue");
const { usePagesPointer } = await import("./usePagesPointer");

afterAll(() => GlobalRegistrator.unregister());

test.each(["header", "footer"] as const)(
  "entering a read-only view preserves pending %s edits and closes its editing surface",
  (position) => {
    const source: Document = {
      package: {
        document: { content: [] },
        headers: new Map(),
        footers: new Map(),
      },
    };
    const parts =
      position === "header" ? source.package.headers : source.package.footers;
    (parts ?? panic("Expected header/footer map")).set("rId-active", {
      type: position,
      content: [],
    });
    let currentDocument = source;
    const hfView = new EditorView(document.createElement("div"), {
      state: EditorState.create({
        doc: schema.node("doc", null, [
          schema.node("paragraph", null, [schema.text("Pending edit")]),
        ]),
      }),
    });
    const readOnly = ref(false);
    const syncHfPMs = mock(() => {});
    const onDocumentChange = mock((_document: Document) => {});
    const holder: { pointer: ReturnType<typeof usePagesPointer> | null } = {
      pointer: null,
    };
    const app = createApp(
      defineComponent({
        setup() {
          holder.pointer = usePagesPointer({
            editorView: shallowRef(null),
            pagesRef: ref(null),
            pagesViewportRef: ref(null),
            scrollRootRef: ref(null),
            selectedImage: shallowRef(null),
            imageInteracting: ref(false),
            hyperlinkPopupData: ref(null),
            readOnly,
            showHeaderFooterEditing: ref(true),
            zoom: ref(1),
            layout: ref(null),
            tableResize: {
              tryStartResize: () => false,
              isResizing: ref(false),
            },
            getCommands: () => ({}),
            getDocument: () => currentDocument,
            getHfPmView: () => hfView,
            reLayout: () => {},
            onDocumentChange,
            clearOverlay: () => {},
            syncHfPMs,
            setDocument: (document) => {
              currentDocument = document;
            },
            openNoteStory: () => {},
            closeNoteStory: () => {},
            getActiveNoteView: () => null,
          });
          return () => h("div");
        },
      }),
    );
    app.mount(document.createElement("div"));
    try {
      const pointer =
        holder.pointer ?? panic("Expected pages pointer composable");
      pointer.hfEdit.value = {
        isFirstPage: false,
        pageNumber: 1,
        position,
        rId: "rId-active",
        targetRect: null,
      };
      readOnly.value = true;
      expect(pointer.hfEdit.value).toBeNull();
      const updatedParts =
        position === "header"
          ? currentDocument.package.headers
          : currentDocument.package.footers;
      expect(updatedParts?.get("rId-active")?.content).toMatchObject([
        {
          type: "paragraph",
          content: [
            { type: "run", content: [{ type: "text", text: "Pending edit" }] },
          ],
        },
      ]);
      expect(parts?.get("rId-active")?.content).toEqual([]);
      expect(onDocumentChange).toHaveBeenCalledTimes(1);
      expect(syncHfPMs).toHaveBeenCalledTimes(1);
      readOnly.value = false;
      readOnly.value = true;
      expect(onDocumentChange).toHaveBeenCalledTimes(1);
    } finally {
      app.unmount();
      hfView.destroy();
    }
  },
);
