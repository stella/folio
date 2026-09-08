import { describe, expect, mock, test } from "bun:test";
import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import * as proseMirrorView from "prosemirror-view";

import { assignParagraphPropertySource } from "../docx/paragraphPropertySource";
import { serializeParagraph } from "../docx/serializer/paragraphSerializer";
import { NAMESPACES } from "../docx/xmlParser";
import type { Document, HeaderFooter, Paragraph } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";

type EditorViewOptions = {
  dispatchTransaction: (transaction: Transaction) => void;
  state: EditorState;
};

class TestEditorView {
  readonly dom: TestElement;
  state: EditorState;
  readonly dispatchTransaction: (transaction: Transaction) => void;

  constructor(dom: TestElement, options: EditorViewOptions) {
    this.dom = dom;
    this.state = options.state;
    this.dispatchTransaction = options.dispatchTransaction;
  }

  destroy(): void {}

  dispatch(transaction: Transaction): void {
    this.dispatchTransaction(transaction);
  }

  updateState(state: EditorState): void {
    this.state = state;
  }
}

class TestElement {
  readonly dataset: Record<string, string> = {};
  readonly ownerDocument: { createElement: () => TestElement };
  hidden = false;
  parentElement: TestElement | null = null;

  constructor() {
    this.ownerDocument = { createElement: () => new TestElement() };
  }

  append(child: TestElement): void {
    child.parentElement = this;
  }

  remove(): void {
    this.parentElement = null;
  }
}

mock.module("prosemirror-view", () => ({
  ...proseMirrorView,
  EditorView: TestEditorView,
}));

const { createHeaderFooterEditorManager } = await import("./headerFooterEditorManager");
const { createNoteEditorManager } = await import("./noteEditorManager");

const paragraphWithSource = (text: string, owner: string, paraId?: string): Paragraph => {
  const paragraph: Paragraph = {
    type: "paragraph",
    ...(paraId ? { paraId } : {}),
    content: [{ type: "run", content: [{ type: "text", text }] }],
  };
  assignParagraphPropertySource(paragraph, {
    formattingJson: canonicalJson({}),
    xml:
      `<w:pPr xmlns:w="${NAMESPACES.w}" xmlns:x="urn:folio:test">` +
      `<x:property x:owner="${owner}"/></w:pPr>`,
  });
  return paragraph;
};

const editParagraphText = (view: TestEditorView, text: string): void => {
  let paragraph: PMNode | undefined;
  let position = -1;
  view.state.doc.descendants((node, pos) => {
    if (paragraph || node.type.name !== "paragraph" || node.textContent !== text) {
      return true;
    }
    paragraph = node;
    position = pos;
    return false;
  });
  if (!paragraph || position < 0) {
    panic(`The test editor has no paragraph containing ${text}.`);
  }
  view.dispatch(view.state.tr.insertText(" edited", position + paragraph.nodeSize - 1));
};

const asHtmlElement = (element: TestElement): HTMLElement => {
  // SAFETY: EditorView is replaced by TestEditorView, and the manager only uses
  // this host's ownerDocument, append, dataset, hidden, parentElement, and remove.
  return element as unknown as HTMLElement;
};

const expectOwner = (paragraph: Paragraph, owner: string): void => {
  const xml = serializeParagraph(paragraph);
  expect(xml).toContain(`x:owner="${owner}"`);
  expect(xml.match(/x:owner=/gu)).toHaveLength(1);
};

describe("secondary-story paragraph property ownership", () => {
  test.each(["header", "footer"] as const)(
    "%s manager snapshots an edited id-less paragraph after allocation and auto-bidi",
    (kind) => {
      const host = new TestElement();
      const relationshipId = `rId-${kind}`;
      const arabic = {
        type: "paragraph",
        content: [{ type: "run", content: [{ type: "text", text: "مرحبا" }] }],
      } satisfies Paragraph;
      const source = paragraphWithSource(`${kind} source`, kind);
      const part = {
        type: kind,
        hdrFtrType: "default",
        content: [arabic, source],
        ...(kind === "header" ? { watermarkBlockIndex: 1 } : {}),
      } satisfies HeaderFooter;
      const document: Document = {
        package: {
          document: { content: [] },
          ...(kind === "header"
            ? { headers: new Map([[relationshipId, part]]) }
            : { footers: new Map([[relationshipId, part]]) }),
        },
      };
      const manager = createHeaderFooterEditorManager({
        getDocument: () => document,
        getHost: () => asHtmlElement(host),
        getStyles: () => undefined,
        getTheme: () => undefined,
      });
      manager.sync();
      const view = manager.getView(relationshipId);
      if (!(view instanceof TestEditorView)) {
        panic("The header/footer manager did not mount its test editor.");
      }
      editParagraphText(view, `${kind} source`);

      const snapshot = manager.snapshotDocument(document);
      const savedPart =
        kind === "header"
          ? snapshot.package.headers?.get(relationshipId)
          : snapshot.package.footers?.get(relationshipId);
      const saved = savedPart?.content.at(1);
      if (saved?.type !== "paragraph") {
        panic("The snapshot lost its edited header/footer paragraph.");
      }
      expect(saved.content).toContainEqual({
        type: "run",
        content: [{ type: "text", text: `${kind} source edited` }],
        formatting: undefined,
      });
      expectOwner(saved, kind);
      manager.destroy();
    },
  );

  test.each(["header", "footer"] as const)(
    "%s manager keeps duplicate-id paragraph captures with their exact owners",
    (kind) => {
      const host = new TestElement();
      const relationshipId = `rId-duplicate-${kind}`;
      const first = paragraphWithSource(`${kind} first`, kind, "11111111");
      const second = paragraphWithSource(`${kind} second`, "other", "11111111");
      const part = {
        type: kind,
        hdrFtrType: "default",
        content: [first, second],
      } satisfies HeaderFooter;
      const document: Document = {
        package: {
          document: { content: [] },
          ...(kind === "header"
            ? { headers: new Map([[relationshipId, part]]) }
            : { footers: new Map([[relationshipId, part]]) }),
        },
      };
      const manager = createHeaderFooterEditorManager({
        getDocument: () => document,
        getHost: () => asHtmlElement(host),
        getStyles: () => undefined,
        getTheme: () => undefined,
      });
      manager.sync();
      const view = manager.getView(relationshipId);
      if (!(view instanceof TestEditorView)) {
        panic("The header/footer manager did not mount its test editor.");
      }
      editParagraphText(view, `${kind} second`);

      const snapshot = manager.snapshotDocument(document);
      const savedPart =
        kind === "header"
          ? snapshot.package.headers?.get(relationshipId)
          : snapshot.package.footers?.get(relationshipId);
      const savedFirst = savedPart?.content.at(0);
      const savedSecond = savedPart?.content.at(1);
      if (savedFirst?.type !== "paragraph" || savedSecond?.type !== "paragraph") {
        panic("The duplicate-id snapshot lost a header/footer paragraph.");
      }
      expectOwner(savedFirst, kind);
      expectOwner(savedSecond, "other");
      manager.destroy();
    },
  );

  test.each(["footnote", "endnote"] as const)(
    "%s manager snapshots an edited id-less paragraph after allocation and auto-bidi",
    (kind) => {
      const host = new TestElement();
      const arabic = {
        type: "paragraph",
        content: [{ type: "run", content: [{ type: "text", text: "مرحبا" }] }],
      } satisfies Paragraph;
      const source = paragraphWithSource(`${kind} source`, kind);
      const document: Document = {
        package: {
          document: { content: [] },
          ...(kind === "footnote"
            ? { footnotes: [{ type: "footnote", id: 1, content: [arabic, source] }] }
            : { endnotes: [{ type: "endnote", id: 1, content: [arabic, source] }] }),
        },
      };
      const manager = createNoteEditorManager({
        getDocument: () => document,
        getHost: () => asHtmlElement(host),
        getStyles: () => undefined,
        getTheme: () => undefined,
      });
      const story = { kind, noteId: 1 } as const;
      const view = manager.activate(story);
      if (!(view instanceof TestEditorView)) {
        panic("The note manager did not mount its test editor.");
      }
      editParagraphText(view, `${kind} source`);

      const snapshot = manager.snapshotDocument(document);
      const savedStory =
        kind === "footnote" ? snapshot.package.footnotes?.at(0) : snapshot.package.endnotes?.at(0);
      const saved = savedStory?.content.at(1);
      if (saved?.type !== "paragraph") {
        panic("The snapshot lost its edited note paragraph.");
      }
      expectOwner(saved, kind);
      manager.destroy();
    },
  );

  test.each(["footnote", "endnote"] as const)(
    "%s manager keeps duplicate-id paragraph captures with their exact owners",
    (kind) => {
      const host = new TestElement();
      const first = paragraphWithSource(`${kind} first`, kind, "11111111");
      const second = paragraphWithSource(`${kind} second`, "other", "11111111");
      const document: Document = {
        package: {
          document: { content: [] },
          ...(kind === "footnote"
            ? { footnotes: [{ type: "footnote", id: 1, content: [first, second] }] }
            : { endnotes: [{ type: "endnote", id: 1, content: [first, second] }] }),
        },
      };
      const manager = createNoteEditorManager({
        getDocument: () => document,
        getHost: () => asHtmlElement(host),
        getStyles: () => undefined,
        getTheme: () => undefined,
      });
      const story = { kind, noteId: 1 } as const;
      const view = manager.activate(story);
      if (!(view instanceof TestEditorView)) {
        panic("The note manager did not mount its test editor.");
      }
      editParagraphText(view, `${kind} second`);

      const snapshot = manager.snapshotDocument(document);
      const savedStory =
        kind === "footnote" ? snapshot.package.footnotes?.at(0) : snapshot.package.endnotes?.at(0);
      const savedFirst = savedStory?.content.at(0);
      const savedSecond = savedStory?.content.at(1);
      if (savedFirst?.type !== "paragraph" || savedSecond?.type !== "paragraph") {
        panic("The duplicate-id snapshot lost a note paragraph.");
      }
      expectOwner(savedFirst, kind);
      expectOwner(savedSecond, "other");
      manager.destroy();
    },
  );
});
