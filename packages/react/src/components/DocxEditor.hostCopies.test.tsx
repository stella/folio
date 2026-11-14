import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

import { panic } from "better-result";
import { afterAll, expect, test } from "bun:test";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { IntlProvider } from "use-intl";

import { sharedObjects } from "@stll/folio-core/__tests__/sharedObjects";
import { cloneDocumentWithParagraphPropertySources } from "@stll/folio-core/docx/document-clone";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import type { LayoutInstrumentation } from "@stll/folio-core/layout-engine/layoutInstrumentation";
import type { Document } from "@stll/folio-core/types/document";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

// The host gets one copy of the editor's document per settled change, in
// both sessions, and never the document the editor holds: a change still in
// progress (typing, a margin drag) copies nothing until it settles.

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const SESSIONS = ["legacy", "canonical"] as const;
type Session = (typeof SESSIONS)[number];

/** Whether `value` is a whole document, not some of its content. */
const isDocument = (value: unknown): boolean =>
  typeof value === "object" &&
  value !== null &&
  "package" in value &&
  typeof value.package === "object" &&
  value.package !== null &&
  "document" in value.package;

/** Counts whole-document copies made from now until `stop` is called. */
const countDocumentCopies = () => {
  const original = globalThis.structuredClone;
  let copies = 0;
  globalThis.structuredClone = ((
    value: unknown,
    options?: StructuredSerializeOptions,
  ) => {
    if (isDocument(value)) copies += 1;
    return original(value, options);
  }) as typeof structuredClone;
  return {
    copies: () => copies,
    stop: () => {
      globalThis.structuredClone = original;
    },
  };
};

const recordChanges = (changes: Document[]) => (changed: Document) => {
  changes.push(changed);
};

const settle = () => new Promise((resolve) => setTimeout(resolve, 600));

const typeInto = (view: EditorView, data: string) =>
  view.dom.dispatchEvent(
    new InputEvent("beforeinput", {
      bubbles: true,
      cancelable: true,
      inputType: "insertText",
      data,
    }),
  );

/** A body paragraph with a footnote reference, and the footnote. */
const documentWithFootnote = (): Document => {
  const source = createEmptyDocument({ initialText: "Body" });
  const paragraph = source.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Expected a body paragraph");
  paragraph.content.push({
    type: "run",
    content: [{ type: "footnoteRef", id: 1 }],
  });
  source.package.footnotes = [
    {
      type: "footnote",
      id: 1,
      content: [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "Note" }] }],
        },
      ],
    },
  ];
  return source;
};

const mountEditor = async (
  session: Session,
  source: Document,
  options: { showRuler?: boolean } = {},
) => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  const changes: Document[] = [];
  const onChange = recordChanges(changes);
  const bytes = await createDocx(source);
  await act(async () =>
    root.render(
      <IntlProvider
        locale="en"
        timeZone="UTC"
        messages={getFolioMessages("en")}
      >
        <DocxEditor
          ref={editor}
          accessibleName="Document editor"
          documentBuffer={bytes}
          experimentalSession={session}
          showToolbar={false}
          {...(options.showRuler ? { showRuler: true } : {})}
          onChange={onChange}
        />
      </IntlProvider>,
    ),
  );
  await act(async () => editor.current?.loadDocumentBuffer(bytes));
  await act(async () => editor.current?.ensureEditorView({ focus: false }));
  await act(settle);
  changes.length = 0;
  const ref = () => editor.current ?? panic("Expected a mounted editor");
  return {
    container,
    changes,
    ref,
    unmount: async () => {
      await act(async () => root.unmount());
      container.remove();
    },
  };
};

/** The host owns what it was handed: emptying it leaves the editor's document. */
const expectHostOwnsCopy = (
  notified: Document,
  editor: () => DocxEditorRef,
) => {
  expect(notified.package.document.content).toEqual(
    editor().getDocument()?.package.document.content ??
      panic("Expected the editor's document"),
  );
  notified.package.document.content.length = 0;
  expect(
    editor().getDocument()?.package.document.content.length,
  ).toBeGreaterThan(0);
};

test.each(SESSIONS)(
  "typing in the %s session's body copies the document once, when the edit settles",
  async (session) => {
    const mounted = await mountEditor(
      session,
      createEmptyDocument({ initialText: "Start" }),
    );
    const counter = countDocumentCopies();
    try {
      const view =
        mounted.ref().getEditorRef()?.getView() ?? panic("Expected a view");
      await act(async () => {
        typeInto(view, "N");
        typeInto(view, "O");
        typeInto(view, "P");
      });
      expect(counter.copies()).toBe(0);
      expect(mounted.changes).toHaveLength(0);
      await act(settle);
      counter.stop();
      expect(view.state.doc.textContent).toContain("NOP");
      expect(mounted.changes).toHaveLength(1);
      expect(counter.copies()).toBe(1);
      expectHostOwnsCopy(
        mounted.changes[0] ?? panic("Expected a notified document"),
        mounted.ref,
      );
    } finally {
      counter.stop();
      await mounted.unmount();
    }
  },
);

test.each(SESSIONS)(
  "typing in a note in the %s session copies the document once, when the edit settles",
  async (session) => {
    const mounted = await mountEditor(session, documentWithFootnote());
    const counter = { stop: () => {} };
    try {
      const paged = mounted.ref().getEditorRef() ?? panic("Expected paged ref");
      const body = paged.getView() ?? panic("Expected a body view");
      // happy-dom paints no pages; this is the delegated double-click path.
      const pages =
        mounted.container.querySelector(".paged-editor__pages") ??
        panic("Expected pages container");
      const target = document.createElement("span");
      target.dataset["noteKind"] = "footnote";
      target.dataset["noteId"] = "1";
      pages.append(target);
      await act(async () =>
        target.dispatchEvent(
          new MouseEvent("click", {
            detail: 2,
            bubbles: true,
            cancelable: true,
          }),
        ),
      );
      const note = paged.getActiveView() ?? panic("Expected the note view");
      expect(note).not.toBe(body);
      await act(async () => {
        note.dispatch(
          note.state.tr.setSelection(
            TextSelection.create(
              note.state.doc,
              note.state.doc.content.size - 1,
            ),
          ),
        );
      });
      await act(settle);
      mounted.changes.length = 0;

      const typing = countDocumentCopies();
      counter.stop = typing.stop;
      await act(async () => {
        typeInto(note, "N");
        typeInto(note, "O");
        typeInto(note, "P");
      });
      expect(typing.copies()).toBe(0);
      expect(mounted.changes).toHaveLength(0);
      await act(settle);
      typing.stop();
      expect(note.state.doc.textContent).toContain("NOP");
      expect(mounted.changes).toHaveLength(1);
      expect(typing.copies()).toBe(1);
      const notified =
        mounted.changes[0] ?? panic("Expected a notified document");
      expect(JSON.stringify(notified.package.footnotes ?? [])).toContain("NOP");
      expectHostOwnsCopy(notified, mounted.ref);
    } finally {
      counter.stop();
      await mounted.unmount();
    }
  },
);

test("a margin drag in the legacy session copies the document once, when the drag settles", async () => {
  const mounted = await mountEditor(
    "legacy",
    createEmptyDocument({ initialText: "Start" }),
    { showRuler: true },
  );
  const counter = { stop: () => {} };
  try {
    const leftMargin =
      mounted.container.querySelector('[aria-label="Horizontal ruler"]')
        ?.firstElementChild ?? panic("Expected the left margin zone");
    const dragging = countDocumentCopies();
    counter.stop = dragging.stop;
    await act(async () =>
      leftMargin.dispatchEvent(
        new MouseEvent("mousedown", { bubbles: true, cancelable: true }),
      ),
    );
    for (const clientX of [60, 70, 80, 90]) {
      await act(async () =>
        document.dispatchEvent(new MouseEvent("mousemove", { clientX })),
      );
    }
    await act(async () => document.dispatchEvent(new MouseEvent("mouseup")));
    expect(dragging.copies()).toBe(0);
    expect(mounted.changes).toHaveLength(0);
    await act(settle);
    dragging.stop();
    expect(mounted.changes).toHaveLength(1);
    expect(dragging.copies()).toBe(1);
    const notified =
      mounted.changes[0] ?? panic("Expected a notified document");
    const margin = notified.package.document.finalSectionProperties?.marginLeft;
    expect(margin).toBe(
      mounted.ref().getDocument()?.package.document.finalSectionProperties
        ?.marginLeft,
    );
    expectHostOwnsCopy(notified, mounted.ref);
  } finally {
    counter.stop();
    await mounted.unmount();
  }
});

test("the notes getDocument hands the host in the legacy session are the host's own", async () => {
  const mounted = await mountEditor("legacy", documentWithFootnote());
  try {
    const handed =
      mounted.ref().getDocument() ?? panic("Expected the editor's document");
    const note = handed.package.footnotes?.at(0) ?? panic("Expected a note");
    expect(JSON.stringify(note.content)).toContain("Note");
    note.content.length = 0;
    handed.package.footnotes?.splice(0);
    const again =
      mounted.ref().getDocument() ?? panic("Expected the editor's document");
    expect(JSON.stringify(again.package.footnotes ?? [])).toContain("Note");
  } finally {
    await mounted.unmount();
  }
});

test("the controller's getDocument in the legacy session is read-only: a host writes its own copy, and the editor's document and its undo stay", async () => {
  const mounted = await mountEditor(
    "legacy",
    createEmptyDocument({ initialText: "Start" }),
  );
  try {
    const view =
      mounted.ref().getEditorRef()?.getView() ?? panic("Expected a view");
    await act(async () => typeInto(view, "N"));
    await act(settle);
    const margin = () =>
      mounted.ref().getDocument()?.package.document.finalSectionProperties
        ?.marginLeft;
    const before = margin();
    expect(before).toBeNumber();

    // The document shares the editor's model, so it is handed out read-only
    // (writing it is a type error): a host that writes copies it first.
    const handed =
      mounted.ref().getEditor()?.getDocument() ??
      panic("Expected the controller's document");
    const owned = cloneDocumentWithParagraphPropertySources(handed);
    const sections =
      owned.package.document.finalSectionProperties ??
      panic("Expected section properties");
    sections.marginLeft = 4321;
    expect(handed.package.document.finalSectionProperties?.marginLeft).toBe(
      before,
    );

    expect(margin()).toBe(before);
    await act(async () => mounted.ref().undo());
    await act(settle);
    expect(view.state.doc.textContent).not.toContain("N");
    expect(margin()).toBe(before);
  } finally {
    await mounted.unmount();
  }
});

/** Counts the copy owner's whole-document copies until `stop` is called. */
const countCopyOwnerCopies = () => {
  const previous: LayoutInstrumentation | undefined =
    globalThis.__folioLayoutInstrumentation;
  let copies = 0;
  globalThis.__folioLayoutInstrumentation = {
    ...previous,
    onDocumentBuilt: (event) => {
      previous?.onDocumentBuilt?.(event);
      if (event.kind === "copy") copies += 1;
    },
  };
  return {
    copies: () => copies,
    stop: () => {
      globalThis.__folioLayoutInstrumentation = previous;
    },
  };
};

test("getDocument in the legacy session takes one copy, and shares no object with the editor's document or an earlier call's", async () => {
  const mounted = await mountEditor("legacy", documentWithFootnote());
  const counter = countCopyOwnerCopies();
  try {
    const view =
      mounted.ref().getEditorRef()?.getView() ?? panic("Expected a view");
    await act(async () => typeInto(view, "N"));
    await act(settle);
    const live =
      mounted.ref().getEditorRef()?.getDocument() ??
      panic("Expected the editor's document");
    const before = counter.copies();
    const handed =
      mounted.ref().getDocument() ?? panic("Expected the editor's document");
    expect(counter.copies() - before).toBe(1);
    expect(JSON.stringify(handed.package.document.content)).toContain("N");
    expect(JSON.stringify(handed.package.footnotes ?? [])).toContain("Note");
    expect(sharedObjects(live, handed)).toEqual([]);
    // Two calls share nothing, so neither shares what both read: the
    // committed document's package.
    const again =
      mounted.ref().getDocument() ?? panic("Expected the editor's document");
    expect(again).toEqual(handed);
    expect(sharedObjects(handed, again)).toEqual([]);
  } finally {
    counter.stop();
    await mounted.unmount();
  }
});
