import { GlobalRegistrator } from "@happy-dom/global-registrator";

// PM keymaps capture the host platform at import time. Keep the DOM fixture on
// that same platform when another test has already imported the editor schema.
const hostPlatform = typeof navigator === "undefined" ? "" : navigator.platform;
GlobalRegistrator.register();
Object.defineProperty(navigator, "platform", { configurable: true, value: hostPlatform });

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { TextSelection } from "prosemirror-state";
import { closeHistory } from "prosemirror-history";

const { createApp, defineComponent, h, ref, shallowRef, nextTick } = await import("vue");

import { parseDocx } from "@stll/folio-core/docx/parser";
import { validateDocxPackage } from "../../../docx-core/src/validate/docx";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { CanonicalSessionError } from "@stll/folio-core/controller/canonicalSession";
import { CanonicalDocxInputError } from "@stll/folio-core/docx/canonicalSessionInput";
import { createEmptyHeaderFooter } from "@stll/folio-core/utils/headerFooter";
import type { Comment } from "@stll/folio-core/types/content";
import { getDocumentWatermark, type Watermark } from "@stll/folio-core/watermark";
import { normalizeCanonicalWatermark } from "../../../docx-core/src/model/watermarkDefaults";

const { isMacPlatform } = await import("@stll/folio-core/managers/editorShortcuts");
import {
  createCanonicalHeaderFooterOperation,
  DOCUMENT_OP_TYPES,
  removeCanonicalHeaderFooterOperations,
  withCanonicalParagraphIds,
} from "@stll/folio-core/controller/canonicalOperations";

const { usePageSetupControls } = await import("./usePageSetupControls");
import { reviewDifferences } from "../../../../test/reviewDifferences";
import { expectCanonicalWatermarkRoundTrip } from "../../../../test/canonicalWatermarkRoundTrip";
import { CanonicalSaveDiagnosticError } from "@stll/folio-core/docx/canonicalSave";
import type { SaveDiagnostic } from "@stll/folio-core/docx/saveDiagnostics";
import { describePackageDifferences } from "../../../../scripts/lib/corpus-invariants/model-equality";
import {
  CANONICAL_SAVE_SEEDS,
  canonicalSaveFixture,
  canonicalSaveSequence,
  canonicalSaveParagraphXml,
} from "../../../../test/canonicalSaveSequence";
import {
  clearTrackedChanges,
  getChangedParagraphIds,
  hasStructuralChanges,
} from "@stll/folio-core/prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { canonicalReviewBlocks } from "../../../../test/reviewProjection";

const { useDocxEditor } = await import("./useDocxEditor");
const { useCommentManagement } = await import("./useCommentManagement");
const { useKeyboardShortcuts } = await import("./useKeyboardShortcuts");

// The save oracle exercises the composable's real hidden manager and serialization,
// rather than rebuilding the expected document from its PM projection.
afterAll(() => GlobalRegistrator.unregister());

test.each([
  {
    kind: "encrypted",
    bytes: new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
    message: "Password-protected documents are unavailable in the experimental canonical session.",
  },
  {
    kind: "malformed",
    bytes: new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
    message: "Failed to normalize paragraph IDs: Failed to parse DOCX archive",
  },
])("canonical $kind loads preserve the typed input error", async ({ bytes, message }) => {
  const container = document.createElement("div");
  document.body.append(container);
  const errors: Error[] = [];
  const holder: { editor: ReturnType<typeof useDocxEditor> | null } = { editor: null };
  const app = createApp(
    defineComponent({
      setup() {
        holder.editor = useDocxEditor({
          hiddenContainer: shallowRef(null),
          pagesContainer: shallowRef(null),
          experimentalSession: "canonical",
          onError: (error) => errors.push(error),
        });
        return () => h("div");
      },
    }),
  );
  app.mount(container);
  try {
    const editor = holder.editor ?? panic("Expected mounted Vue editor");
    await editor.loadBuffer(bytes);
    expect(errors).toHaveLength(1);
    expect(errors.at(0)).toBeInstanceOf(CanonicalDocxInputError);
    expect(errors.at(0)?.message).toBe(message);
    expect(editor.parseError.value).toBe(message);
    expect(editor.isReady.value).toBe(false);
    expect(editor.getDocument()).toBeNull();
  } finally {
    app.unmount();
    container.remove();
  }
});
test.each(["canonical", "default"] as const)(
  "%s history keys have one editor owner and no document-level owner",
  async (session) => {
    const container = document.createElement("div");
    const hidden = document.createElement("div");
    const pages = document.createElement("div");
    document.body.append(container, hidden, pages);
    const holder: { editor: ReturnType<typeof useDocxEditor> | null } = { editor: null };
    const app = createApp(
      defineComponent({
        setup() {
          holder.editor = useDocxEditor({
            hiddenContainer: shallowRef(hidden),
            pagesContainer: shallowRef(pages),
            ...(session === "canonical" ? { experimentalSession: "canonical" as const } : {}),
          });
          useKeyboardShortcuts({
            showFindReplace: shallowRef(false),
            showHyperlink: shallowRef(false),
            handleZoomKeyDown: () => {},
            scope: () => "document",
            roots: [shallowRef(container)],
          });
          return () => h("div");
        },
      }),
    );
    app.mount(container);
    try {
      const editor = holder.editor ?? panic("Expected mounted Vue editor");
      await editor.loadBuffer(await createDocx(createEmptyDocument({ initialText: "Start" })));
      const view = editor.editorView.value ?? panic("Expected body editor view");
      const insert = (text: string) => {
        view.dispatch(
          closeHistory(
            view.state.tr.setSelection(
              TextSelection.create(view.state.doc, view.state.doc.content.size - 1),
            ),
          ),
        );
        if (session === "canonical") {
          view.dom.dispatchEvent(
            new InputEvent("beforeinput", {
              bubbles: true,
              cancelable: true,
              inputType: "insertText",
              data: text,
            }),
          );
        } else {
          view.dispatch(view.state.tr.insertText(text));
        }
      };
      insert("1");
      insert("2");
      const key = (chord: string) =>
        new KeyboardEvent("keydown", {
          bubbles: true,
          cancelable: true,
          ctrlKey: !isMacPlatform(),
          metaKey: isMacPlatform(),
          key: chord,
        });
      for (const chord of ["z", "y"]) {
        const event = key(chord);
        document.dispatchEvent(event);
        expect(event.defaultPrevented).toBe(false);
        expect(view.state.doc.textContent).toBe("Start12");
      }
      const undo = key("z");
      view.dom.dispatchEvent(undo);
      expect(undo.defaultPrevented).toBe(true);
      expect(view.state.doc.textContent).toBe("Start1");
      const redo = key("y");
      view.dom.dispatchEvent(redo);
      expect(redo.defaultPrevented).toBe(true);
      expect(view.state.doc.textContent).toBe("Start12");
    } finally {
      app.unmount();
      container.remove();
      hidden.remove();
      pages.remove();
    }
  },
);

test("canonical edits save and reopen the canonical text and paragraph identity", async () => {
  const container = document.createElement("div");
  const hidden = document.createElement("div");
  const pages = document.createElement("div");
  document.body.append(container, hidden, pages);
  const errors: Error[] = [];
  let hostChanges = 0;
  let hostInputs = 0;
  const holder: { editor: ReturnType<typeof import("./useDocxEditor").useDocxEditor> | null } = {
    editor: null,
  };
  const app = createApp(
    defineComponent({
      setup() {
        holder.editor = useDocxEditor({
          hiddenContainer: shallowRef(hidden),
          pagesContainer: shallowRef(pages),
          experimentalSession: "canonical",
          onError: (error) => errors.push(error),
          onChange: (document) => {
            hostChanges++;
            document.package.document.content = [];
          },
        });
        return () =>
          h("div", {
            onInput: () => {
              hostInputs++;
            },
          });
      },
    }),
  );
  app.mount(container);
  try {
    const rendered = container.firstElementChild ?? panic("Expected mounted Vue DOM");
    expect(rendered.ownerDocument).toBe(document);
    expect(rendered instanceof HTMLElement).toBe(true);
    rendered.dispatchEvent(new InputEvent("input", { bubbles: true }));
    expect(hostInputs).toBe(1);
    const editor = holder.editor ?? panic("Expected mounted Vue editor");
    const bytes = await createDocx(createEmptyDocument({ initialText: "Start" }));
    await editor.loadBuffer(bytes);
    const detachedRead = editor.getDocument() ?? panic("Expected canonical document snapshot");
    detachedRead.package.document.content = [];
    expect(editor.getDocument()?.package.document.content).not.toHaveLength(0);
    const unsupportedWrite = editor.getDocument() ?? panic("Expected canonical document snapshot");
    unsupportedWrite.package.document.content = [];
    editor.setDocument(unsupportedWrite);
    expect(editor.getDocument()?.package.document.content).not.toHaveLength(0);
    expect(errors.at(0)?.message).toContain("Direct document model changes are unavailable");
    expect(editor.parseError.value).toBeNull();
    expect(editor.isReady.value).toBe(true);
    const view = editor.editorView.value ?? panic("Expected body editor view");
    // Exercise the hidden manager's refusal callback as well as model writes.
    const bold = view.state.schema.marks["bold"] ?? panic("Expected bold mark");
    view.dispatch(view.state.tr.addMark(1, 6, bold.create()));
    expect(errors.at(-1)?.message).toContain("unavailable");
    expect(editor.parseError.value).toBeNull();
    expect(editor.isReady.value).toBe(true);
    expect(view.state.doc.rangeHasMark(1, 6, bold)).toBe(false);
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 6)));
    view.dom.dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertText",
        data: " edited",
      }),
    );
    const canonical = editor.getDocument() ?? panic("Expected canonical document");
    expect(view.state.doc.textContent).toBe("Start edited");
    expect(editor.isDirty.value).toBe(true);
    await new Promise((resolve) => window.setTimeout(resolve, 300));
    expect(hostChanges).toBeGreaterThan(0);
    const saved = await editor.save();
    expect(errors).toHaveLength(2);
    expect(editor.parseError.value).toBeNull();
    if (!saved) panic("Expected saved canonical DOCX");
    const reopened = await parseDocx(await saved.arrayBuffer(), {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(reviewDifferences(canonical, reopened)).toEqual({ messages: [], omitted: 0 });
    expect(editor.isDirty.value).toBe(false);
  } finally {
    app.unmount();
    container.remove();
    hidden.remove();
    pages.remove();
  }
});

test("Vue canonical comment projection follows controlled edits, undo, callbacks, and save", async () => {
  const container = document.createElement("div");
  const hidden = document.createElement("div");
  const pages = document.createElement("div");
  document.body.append(container, hidden, pages);
  const bytes = await createDocx(createEmptyDocument({ initialText: "Comment anchor" }));
  const holder: {
    editor: ReturnType<typeof useDocxEditor> | null;
    comments: ReturnType<typeof useCommentManagement> | null;
  } = { editor: null, comments: null };
  const controlledComments = ref<Comment[] | undefined>([]);
  const stateTick = ref(0);
  const changes: Comment[][] = [];
  const app = createApp(
    defineComponent({
      setup() {
        const editor = useDocxEditor({
          hiddenContainer: shallowRef(hidden),
          pagesContainer: shallowRef(pages),
          experimentalSession: "canonical",
        });
        holder.editor = editor;
        editor.editor.on("docChange", () => {
          stateTick.value += 1;
        });
        holder.comments = useCommentManagement({
          editor: editor.editor,
          editorView: editor.editorView,
          getDocument: editor.getDocument,
          author: () => "Reviewer",
          commentsProp: () => controlledComments.value,
          canonicalTick: stateTick,
          onCommentsChange: (next) => {
            changes.push(structuredClone(next));
            controlledComments.value = structuredClone(next);
          },
          reLayout: () => undefined,
        });
        return () => h("div");
      },
    }),
  );
  app.mount(container);
  try {
    const adapter = holder.editor ?? panic("Expected Vue editor");
    const management = holder.comments ?? panic("Expected Vue comment management");
    await adapter.loadBuffer(bytes);
    stateTick.value += 1;
    await nextTick();
    const created = adapter.editor.applyCanonicalComment({
      type: "create",
      text: "Review this",
      author: "Reviewer",
      anchor: { kind: "selection", from: 1, to: 8, story: "main" },
    });
    expect(created?.status).toBe("applied");
    if (created?.status !== "applied" || created.commentId === undefined) {
      panic("Expected canonical comment creation");
    }
    await new Promise((resolve) => setTimeout(resolve, 275));
    await nextTick();
    expect(management.comments.value).toEqual(created.comments);
    expect(changes.at(-1)).toEqual(created.comments);

    const revisedContent: Comment["content"] = [
      {
        type: "paragraph",
        paraId:
          created.comments[0]?.content[0]?.type === "paragraph"
            ? created.comments[0].content[0].paraId
            : undefined,
        formatting: {},
        content: [{ type: "run", formatting: {}, content: [{ type: "text", text: "Updated" }] }],
      },
    ];
    controlledComments.value = created.comments.map((comment) =>
      comment.id === created.commentId
        ? { ...comment, done: true, content: revisedContent }
        : comment,
    );
    await nextTick();
    await new Promise((resolve) => setTimeout(resolve, 275));
    await nextTick();
    expect(adapter.getDocument()?.package.document.comments).toEqual(controlledComments.value);

    expect(adapter.editor.undo()).toBe(true);
    await new Promise((resolve) => setTimeout(resolve, 275));
    await nextTick();
    const undone = adapter.getDocument()?.package.document.comments ?? [];
    expect(undone.at(0)?.done).toBeFalsy();
    expect(undone.at(0)?.content).toEqual(created.comments.at(0)?.content);
    expect(controlledComments.value).toEqual(undone);
    expect(changes.at(-1)).toEqual(undone);

    const view = adapter.editorView.value ?? panic("Expected Vue canonical view");
    const committedDoc = view.state.doc;
    view.dom.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true }));
    expect(adapter.editor.getCanonicalDocument).toThrow(
      "Composition must finish before taking a snapshot.",
    );
    // Every selection tick recomputes the comment projection while IME blocks snapshots.
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
    stateTick.value += 1;
    controlledComments.value = [...undone];
    await nextTick();
    expect(management.comments.value).toEqual(undone);
    expect(adapter.editor.getCanonicalDocument).toThrow(
      "Composition must finish before taking a snapshot.",
    );
    view.dom.dispatchEvent(new CompositionEvent("compositionend", { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 40));
    await nextTick();
    expect(view.state.doc.eq(committedDoc)).toBe(true);
    expect(adapter.editor.getCanonicalDocument()?.package.document.comments).toEqual(undone);

    const expected =
      adapter.editor.getCanonicalDocument() ?? panic("Expected canonical comments model");
    for (const selective of [false, true]) {
      const saved = await adapter.save({ selective });
      if (!saved) panic("Expected saved canonical comments");
      const buffer = await saved.arrayBuffer();
      expect(await validateDocxPackage(new Uint8Array(buffer))).toEqual({ valid: true });
      const reopened = await parseDocx(buffer, { preloadFonts: false, detectVariables: false });
      expect(reopened.package.document.comments).toEqual(undone);
      expect(describePackageDifferences(expected, reopened)).toEqual({ messages: [], omitted: 0 });
    }
  } finally {
    app.unmount();
    container.remove();
    hidden.remove();
    pages.remove();
  }
});

test("Vue canonical stories share history and save headers, first-page footer, notes and section properties", async () => {
  const container = document.createElement("div");
  const hidden = document.createElement("div");
  const pages = document.createElement("div");
  const notes = document.createElement("div");
  document.body.append(container, hidden, pages, notes);
  const errors: Error[] = [];
  const holder: { editor: ReturnType<typeof useDocxEditor> | null } = { editor: null };
  const app = createApp(
    defineComponent({
      setup() {
        holder.editor = useDocxEditor({
          hiddenContainer: shallowRef(hidden),
          pagesContainer: shallowRef(pages),
          noteEditorContainer: shallowRef(notes),
          experimentalSession: "canonical",
          featureFlags: () => ({ selectiveSave: true }),
          onError: (error) => errors.push(error),
        });
        return () => h("div");
      },
    }),
  );
  app.mount(container);
  try {
    const adapter = holder.editor ?? panic("Expected mounted Vue editor");
    const source =
      createEmptyHeaderFooter(createEmptyDocument({ initialText: "Body" }), "header", false) ??
      panic("Expected document with header");
    const rId =
      [...(source.package.headers?.keys() ?? [])].at(0) ?? panic("Expected header identity");
    await adapter.loadBuffer(await createDocx(source));
    const header = adapter.getHeaderFooterView(rId) ?? panic("Expected canonical header view");
    expect(
      header.state.doc.eq(
        adapter.editor.getCanonicalStoryProjection({ kind: "header", rId }) ??
          panic("Expected header projection"),
      ),
    ).toBe(true);
    header.dom.dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertText",
        data: "Header",
      }),
    );
    expect(errors).toEqual([]);
    expect(header.state.doc.textContent).toBe("Header");
    expect(adapter.isDirty.value).toBe(true);
    expect(
      adapter.editor.applyCanonicalStoryHistory({
        view: header,
        story: { kind: "header", rId },
        direction: "undo",
      }),
    ).toBe(true);
    expect(header.state.doc.textContent).toBe("");
    expect(
      adapter.editor.applyCanonicalStoryHistory({
        view: header,
        story: { kind: "header", rId },
        direction: "redo",
      }),
    ).toBe(true);
    expect(header.state.doc.textContent).toBe("Header");

    const beforeFooter = adapter.getDocument() ?? panic("Expected canonical document");
    const footerOp = createCanonicalHeaderFooterOperation({
      document: beforeFooter,
      position: "footer",
      referenceType: "first",
    });
    expect(adapter.editor.applyCanonicalOperations([footerOp])).toBe(true);
    const footer =
      adapter.getHeaderFooterView(footerOp.story.rId) ?? panic("Expected first-page footer view");
    footer.dom.dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertText",
        data: "First footer",
      }),
    );
    expect(footer.state.doc.textContent).toBe("First footer");

    const beforeNote = adapter.getDocument() ?? panic("Expected canonical document");
    const paragraph = beforeNote.package.document.content.at(0);
    if (paragraph?.type !== "paragraph" || !paragraph.paraId)
      panic("Expected addressed body paragraph");
    expect(
      adapter.editor.applyCanonicalOperations([
        {
          type: DOCUMENT_OP_TYPES.ADD_NOTE,
          at: { story: "main", blockId: paragraph.paraId, offset: 4 },
          note: {
            type: "footnote",
            id: 1,
            content: withCanonicalParagraphIds([{ type: "paragraph", content: [] }], beforeNote),
          },
        },
      ]),
    ).toBe(true);
    adapter.openNoteStory({ kind: "footnote", noteId: 1 });
    const note = adapter.getActiveNoteView() ?? panic("Expected canonical footnote view");
    note.dom.dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertText",
        data: "Note",
      }),
    );
    expect(note.state.doc.textContent).toBe("Note");
    expect(
      adapter.editor.applyCanonicalStoryHistory({
        view: note,
        story: { kind: "footnote", id: 1 },
        direction: "undo",
      }),
    ).toBe(true);
    expect(note.state.doc.textContent).toBe("");
    expect(
      adapter.editor.applyCanonicalStoryHistory({
        view: note,
        story: { kind: "footnote", id: 1 },
        direction: "redo",
      }),
    ).toBe(true);
    expect(note.state.doc.textContent).toBe("Note");

    const controls = usePageSetupControls({
      editorView: adapter.editorView,
      getDocument: adapter.getDocument,
      readOnly: shallowRef(false),
      stateTick: shallowRef(0),
      reLayout: adapter.reLayout,
      onChange: () => panic("Canonical section writes must use the journal"),
      applySectionProperties: (properties) => {
        const result = adapter.editor.applyCanonicalSectionProperties(properties);
        if (result === null) return "unhandled";
        return result.status;
      },
    });
    controls.handlePageSetupApply({ marginLeft: 720, footnotePr: { numStart: 2 } });
    const sectionCanonical = adapter.getDocument() ?? panic("Expected canonical document");
    expect(sectionCanonical.package.document.finalSectionProperties?.marginLeft).toBe(720);
    expect(sectionCanonical.package.document.finalSectionProperties?.titlePg).toBe(true);
    expect(adapter.editor.undo()).toBe(true);
    expect(adapter.getDocument()?.package.document.finalSectionProperties?.marginLeft).not.toBe(
      720,
    );
    expect(adapter.editor.redo()).toBe(true);
    expect(adapter.getDocument()?.package.document.finalSectionProperties?.marginLeft).toBe(720);

    const assertWatermarkSave = async () => {
      const canonical = adapter.getDocument() ?? panic("Expected canonical watermark model");
      for (const selective of [false, true]) {
        const saved = await adapter.save({ selective });
        if (!saved) panic("Expected saved canonical watermark");
        const buffer = await saved.arrayBuffer();
        expect(await validateDocxPackage(new Uint8Array(buffer))).toEqual({ valid: true });
        expectCanonicalWatermarkRoundTrip(
          canonical,
          await parseDocx(buffer, { preloadFonts: false, detectVariables: false }),
        );
      }
    };
    const textWatermark = {
      kind: "text",
      text: "DRAFT",
      font: "Calibri",
      color: "C0C0C0",
      diagonal: true,
    } satisfies Watermark;
    expect(
      adapter.editor.applyCanonicalWatermark({ kind: "set", watermark: textWatermark })?.status,
    ).toBe("applied");
    expect(
      getDocumentWatermark(adapter.getDocument() ?? panic("Expected canonical document")),
    ).toEqual(textWatermark);
    expect(adapter.editor.undo()).toBe(true);
    expect(
      getDocumentWatermark(adapter.getDocument() ?? panic("Expected canonical document")),
    ).toBe(undefined);
    expect(adapter.editor.redo()).toBe(true);
    expect(
      getDocumentWatermark(adapter.getDocument() ?? panic("Expected canonical document")),
    ).toEqual(textWatermark);

    await assertWatermarkSave();

    const pictureRequest = {
      kind: "picture",
      imageRId: "rIdExternalWatermark",
      imageTarget: "https://example.test/watermark.png",
      imageTargetExternal: true,
      scale: 0.6,
      widthPt: 249,
      heightPt: 124.2,
    } satisfies Watermark;
    const pictureWatermark = normalizeCanonicalWatermark(pictureRequest);
    expect(
      adapter.editor.applyCanonicalWatermark({ kind: "set", watermark: pictureRequest })?.status,
    ).toBe("applied");
    expect(
      getDocumentWatermark(adapter.getDocument() ?? panic("Expected canonical document")),
    ).toEqual(pictureWatermark);
    expect(adapter.editor.undo()).toBe(true);
    expect(
      getDocumentWatermark(adapter.getDocument() ?? panic("Expected canonical document")),
    ).toEqual(textWatermark);
    expect(adapter.editor.redo()).toBe(true);
    const pictureCanonical = adapter.getDocument() ?? panic("Expected canonical document");
    expect(getDocumentWatermark(pictureCanonical)).toEqual(pictureWatermark);
    await assertWatermarkSave();

    expect(adapter.editor.applyCanonicalWatermark({ kind: "remove" })?.status).toBe("applied");
    expect(
      getDocumentWatermark(adapter.getDocument() ?? panic("Expected canonical document")),
    ).toBe(undefined);
    expect(adapter.editor.undo()).toBe(true);
    expect(
      getDocumentWatermark(adapter.getDocument() ?? panic("Expected canonical document")),
    ).toEqual(pictureWatermark);
    await assertWatermarkSave();
    expect(adapter.editor.redo()).toBe(true);
    const canonical = adapter.getDocument() ?? panic("Expected canonical document");
    expect(getDocumentWatermark(canonical)).toBeUndefined();
    for (const selective of [false, true]) {
      const saved = await adapter.save({ selective });
      if (!saved) panic("Expected canonical story save");
      const buffer = await saved.arrayBuffer();
      expect(await validateDocxPackage(new Uint8Array(buffer))).toEqual({ valid: true });
      const reopened = await parseDocx(buffer, {
        preloadFonts: false,
        detectVariables: false,
      });
      expectCanonicalWatermarkRoundTrip(canonical, reopened);
      expect(reopened.package.document.finalSectionProperties).toEqual(
        canonical.package.document.finalSectionProperties,
      );
      expect([...(reopened.package.headers?.keys() ?? [])].sort()).toEqual(
        [...(canonical.package.headers?.keys() ?? [])].sort(),
      );
      expect([...(reopened.package.footers?.keys() ?? [])].sort()).toEqual(
        [...(canonical.package.footers?.keys() ?? [])].sort(),
      );
      for (const [identity, part] of canonical.package.headers ?? []) {
        expect(
          canonicalReviewBlocks(reopened.package.headers?.get(identity)?.content ?? []),
        ).toEqual(canonicalReviewBlocks(part.content));
      }
      for (const [identity, part] of canonical.package.footers ?? []) {
        expect(
          canonicalReviewBlocks(reopened.package.footers?.get(identity)?.content ?? []),
        ).toEqual(canonicalReviewBlocks(part.content));
      }
      expect((reopened.package.footnotes ?? []).map(({ id }) => id)).toEqual(
        (canonical.package.footnotes ?? []).map(({ id }) => id),
      );
      for (const part of canonical.package.footnotes ?? []) {
        expect(
          canonicalReviewBlocks(
            reopened.package.footnotes?.find(({ id }) => id === part.id)?.content ?? [],
          ),
        ).toEqual(canonicalReviewBlocks(part.content));
      }
    }
    expect(adapter.isDirty.value).toBe(false);
    expect(
      adapter.editor.applyCanonicalOperations(
        removeCanonicalHeaderFooterOperations({ document: canonical, position: "header", rId }),
      ),
    ).toBe(true);
    const headerRemovedCanonical = adapter.getDocument() ?? panic("Expected removed header model");
    const withoutHeader = await adapter.save();
    if (!withoutHeader) panic("Expected saved header removal");
    const removed = await parseDocx(await withoutHeader.arrayBuffer(), {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(removed.package.headers?.has(rId) ?? false).toBe(false);
    expect(removed.package.document.finalSectionProperties?.headerReferences ?? []).toEqual(
      headerRemovedCanonical.package.document.finalSectionProperties?.headerReferences ?? [],
    );
    expect(errors).toHaveLength(5);
    for (const error of errors) {
      expect(error).toBeInstanceOf(CanonicalSaveDiagnosticError);
      if (!(error instanceof CanonicalSaveDiagnosticError)) panic("Expected typed save diagnostic");
      expect(error.gap).toBe("pm-save-projection");
      expect(error.diagnostic).toEqual({ type: "selectiveSaveRefused", part: "word/document.xml" });
    }
  } finally {
    app.unmount();
    container.remove();
    hidden.remove();
    pages.remove();
    notes.remove();
  }
});

test.each(CANONICAL_SAVE_SEEDS)(
  "generated canonical history %s saves independently of PM trackers",
  async (seed) => {
    const container = document.createElement("div");
    const hidden = document.createElement("div");
    const pages = document.createElement("div");
    document.body.append(container, hidden, pages);
    const errors: Error[] = [];
    const flags = { selectiveSave: true, selectiveSaveMaxBytes: seed % 2 === 0 ? undefined : 1 };
    let callbackCount = 0;
    const holder: { editor: ReturnType<typeof useDocxEditor> | null } = { editor: null };
    const app = createApp(
      defineComponent({
        setup() {
          holder.editor = useDocxEditor({
            hiddenContainer: shallowRef(hidden),
            pagesContainer: shallowRef(pages),
            experimentalSession: "canonical",
            featureFlags: () => flags,
            onError: (error) => errors.push(error),
          });
          return () => h("div");
        },
      }),
    );
    app.mount(container);
    try {
      const adapter = holder.editor ?? panic("Expected Vue adapter");
      const bytes = await createDocx(canonicalSaveFixture(seed));
      await adapter.loadBuffer(bytes);
      const api = adapter.editor;
      const view = adapter.editorView.value ?? panic("Expected canonical body view");
      const untouchedId = (seed + 3).toString(16).padStart(8, "0").toUpperCase();
      const baselineParagraphs = new Map<string, string>();
      for (let index = 0; index < 4; index++) {
        const paraId = (seed + index).toString(16).padStart(8, "0").toUpperCase();
        baselineParagraphs.set(paraId, await canonicalSaveParagraphXml(bytes, paraId));
      }
      for (const [index, operations] of canonicalSaveSequence(seed).entries()) {
        const before = adapter.getDocument();
        expect(api.applyCanonicalOperations(operations)).toBe(true);
        const canonical = adapter.getDocument() ?? panic("Expected canonical snapshot");
        expect(api.undo()).toBe(true);
        expect(adapter.getDocument()).toEqual(before);
        expect(api.redo()).toBe(true);
        expect(adapter.getDocument()).toEqual(canonical);
        view.dispatch(clearTrackedChanges(view.state));
        expect(getChangedParagraphIds(view.state).size).toBe(0);
        const changedIds = new Set(
          api.captureCanonicalSave()?.changedBlockIds ?? panic("Expected canonical save capture"),
        );
        expect(changedIds.has(untouchedId)).toBe(false);
        expect(hasStructuralChanges(view.state)).toBe(false);
        const saved = await adapter.save({ selective: index % 2 === 0 });
        if (!saved) panic("Expected generated canonical save");
        const savedBytes = await saved.arrayBuffer();
        const reopened = await parseDocx(savedBytes, {
          preloadFonts: false,
          detectVariables: false,
        });
        expect(describePackageDifferences(canonical, reopened)).toEqual({
          messages: [],
          omitted: 0,
        });
        for (const [paraId, originalXml] of baselineParagraphs) {
          if (!changedIds.has(paraId))
            expect(await canonicalSaveParagraphXml(savedBytes, paraId)).toBe(originalXml);
        }
        expect(api.undo()).toBe(true);
        expect(adapter.getDocument()).toEqual(before);
        expect(api.redo()).toBe(true);
        const diagnostics: SaveDiagnostic[] = [];
        const errorsBeforeSerialization = errors.length;
        const repeated = await api.getDocx({
          mode: index % 2 === 0 ? "full" : "prefer-selective",
          onDiagnostic: (diagnostic) => diagnostics.push(diagnostic),
        });
        expect(errors.length).toBe(errorsBeforeSerialization);
        callbackCount += diagnostics.length;
        for (const diagnostic of diagnostics) {
          expect(diagnostic).toEqual({ type: "selectiveSaveRefused", part: "word/document.xml" });
        }
        if (!repeated) panic("Expected repeated canonical serialization");
        expect(
          describePackageDifferences(
            canonical,
            await parseDocx(repeated, { preloadFonts: false, detectVariables: false }),
          ),
        ).toEqual({ messages: [], omitted: 0 });
        for (const [paraId, originalXml] of baselineParagraphs) {
          if (!changedIds.has(paraId))
            expect(await canonicalSaveParagraphXml(repeated, paraId)).toBe(originalXml);
        }
      }
      const captured = api.captureCanonicalSave() ?? panic("Expected save snapshot");
      const pendingSave = adapter.save({ selective: false });
      expect(api.updateCanonicalInputLifecycle("beginComposition")).toBe(true);
      expect(api.isCanonicalSaveCurrent(captured.version)).toBe(false);
      const capturedSave = await pendingSave;
      if (!capturedSave) panic("Expected captured save during later composition");
      const capturedBytes = await capturedSave.arrayBuffer();
      expect(
        describePackageDifferences(
          captured.document,
          await parseDocx(capturedBytes, {
            preloadFonts: false,
            detectVariables: false,
          }),
        ),
      ).toEqual({ messages: [], omitted: 0 });
      expect(api.updateCanonicalInputLifecycle("endComposition")).toBe(true);
      if (seed % 2 !== 0) {
        expect(errors.length).toBeGreaterThan(0);
        expect(callbackCount).toBeGreaterThan(0);
      }
      for (const error of errors) {
        expect(error).toBeInstanceOf(CanonicalSaveDiagnosticError);
        if (!(error instanceof CanonicalSaveDiagnosticError))
          panic("Expected typed save diagnostic");
        expect(error.diagnostic).toEqual({
          type: "selectiveSaveRefused",
          part: "word/document.xml",
        });
        expect(error.gap).toBe("pm-save-projection");
      }
      // The prior sequence began composition after capture; also cover capture during composition.
      const errorsBeforeComposition = errors.length;
      expect(api.updateCanonicalInputLifecycle("beginComposition")).toBe(true);
      expect(await adapter.save()).toBeNull();
      expect(await api.getDocx()).toBeNull();
      const compositionErrors = errors.slice(errorsBeforeComposition);
      expect(compositionErrors).toHaveLength(2);
      for (const error of compositionErrors) {
        expect(error).toBeInstanceOf(CanonicalSessionError);
        if (!(error instanceof CanonicalSessionError)) panic("Expected composition save refusal");
        expect(error.gap).toBe("pm-save-projection");
        expect(error.reason).toBe("refused");
      }
      expect(api.updateCanonicalInputLifecycle("endComposition")).toBe(true);
      expect(api.captureCanonicalSave()?.document).toEqual(captured.document);
    } finally {
      app.unmount();
      container.remove();
      hidden.remove();
      pages.remove();
    }
  },
);
