import { panic } from "better-result";
import { expect, test } from "bun:test";
import { TextSelection } from "prosemirror-state";

import { DOCUMENT_OP_TYPES } from "@stll/docx-core/ops";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { getStatedDocumentNumbering } from "../prosemirror/plugins/documentNumbering";
import type { Document, Paragraph } from "../types/document";
import { createHeaderFooterEditorManager } from "./headerFooterEditorManager";
import { createHiddenEditorManager } from "./hiddenEditorManager";
import { createNoteEditorManager } from "./noteEditorManager";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

// A public operation on one paragraph of an open header leaves the caret the
// person put in another paragraph of it where it was.
test("a public operation on an open header keeps its editor's caret", () => {
  const releaseDom = acquireHarnessDom();
  const bodyHost = document.createElement("div");
  const storyHost = document.createElement("div");
  document.body.append(bodyHost, storyHost);
  const source: Document = {
    package: { document: { content: [paragraph("12345678", "Body")] } },
  };
  const manager = createHiddenEditorManager({
    getHost: () => bodyHost,
    getDocument: () => source,
    getDocumentContext: () => source,
    getExperimentalSession: () => "canonical",
    getStyles: () => null,
    getExtensionManager: () => undefined,
    getExternalPlugins: () => [],
    getCollaboration: () => undefined,
    getCollaborationModules: () => null,
    getPrecomputedInitialState: () => null,
    getReadOnly: () => false,
    getDocumentIdentity: () => "story-public-selection",
    onTransaction: () => {
      notes.sync();
      parts.sync();
    },
    onCanonicalCompositionEnd: () => {},
    onSelectionChange: () => {},
    onKeyDown: () => false,
    onCopy: () => {},
    onCut: () => {},
    onPaste: () => {},
    onReadOnlyEditAttempt: () => {},
    onEditorViewReady: () => {},
    onEditorViewDestroy: () => {},
    onRemoteSelectionsChange: () => {},
    onSessionRefusal: (message) => panic(message),
  });
  const deps = {
    getHost: () => storyHost,
    getDocument: () => source,
    getCanonicalApi: () => manager.api,
    readCanonicalDocument: () => manager.readCanonicalDocument(),
    getExperimentalSession: () => "canonical" as const,
    getStyles: () => null,
    getTheme: () => null,
  };
  const notes = createNoteEditorManager(deps);
  const parts = createHeaderFooterEditorManager(deps);
  try {
    manager.ensureView();
    expect(
      manager.api.applyCanonicalOperations([
        {
          type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
          sectionIndex: 0,
          story: { kind: "header", rId: "rIdHeader" },
          referenceType: "default",
          content: [
            paragraph("23456781", "First line"),
            paragraph("23456782", "Second line"),
          ],
        },
      ]),
    ).toBe(true);
    const header = parts.getView("rIdHeader") ?? panic("No header editor");
    // The caret three characters into the header's second paragraph.
    let second = -1;
    header.state.doc.forEach((node, offset) => {
      if (node.textContent === "Second line") second = offset + 1;
    });
    header.dispatch(
      header.state.tr.setSelection(
        TextSelection.create(header.state.doc, second + 3),
      ),
    );
    const projection =
      manager.api.getCanonicalStoryProjection({
        kind: "header",
        rId: "rIdHeader",
      }) ?? panic("No header projection");
    const result = manager.api.applyCanonicalDocumentOperations({
      undefinedReferences: "refuse",
      story: { type: "header", relationshipId: "rIdHeader" },
      snapshot: createFolioAIEditSnapshot(projection),
      batch: {
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "op",
            type: "replaceBlock",
            blockId: "23456781",
            text: "Replaced first line",
          },
        ],
      },
    });
    expect(result?.skipped).toEqual([]);
    parts.sync();
    const view = parts.getView("rIdHeader") ?? panic("No header editor");
    expect(view.state.doc.textContent).toBe("Replaced first lineSecond line");
    const { $head } = view.state.selection;
    expect($head.parent.textContent).toBe("Second line");
    expect($head.parentOffset).toBe(3);
  } finally {
    notes.destroy();
    parts.destroy();
    manager.destroyView();
    bodyHost.remove();
    storyHost.remove();
    releaseDom();
  }
});

// A list a public operation starts is one the editor showing the story knows,
// so its list commands and list state find it.
test("a list a public operation starts reaches the body editor's numbering", () => {
  const releaseDom = acquireHarnessDom();
  const bodyHost = document.createElement("div");
  document.body.append(bodyHost);
  const source: Document = {
    package: { document: { content: [paragraph("12345678", "Body")] } },
  };
  const manager = createHiddenEditorManager({
    getHost: () => bodyHost,
    getDocument: () => source,
    getDocumentContext: () => source,
    getExperimentalSession: () => "canonical",
    getStyles: () => null,
    getExtensionManager: () => undefined,
    getExternalPlugins: () => [],
    getCollaboration: () => undefined,
    getCollaborationModules: () => null,
    getPrecomputedInitialState: () => null,
    getReadOnly: () => false,
    getDocumentIdentity: () => "public-new-list-numbering",
    onTransaction: () => {},
    onCanonicalCompositionEnd: () => {},
    onSelectionChange: () => {},
    onKeyDown: () => false,
    onCopy: () => {},
    onCut: () => {},
    onPaste: () => {},
    onReadOnlyEditAttempt: () => {},
    onEditorViewReady: () => {},
    onEditorViewDestroy: () => {},
    onRemoteSelectionsChange: () => {},
    onSessionRefusal: (message) => panic(message),
  });
  try {
    manager.ensureView();
    const view = manager.getView() ?? panic("No body view");
    const result = manager.api.applyCanonicalDocumentOperations({
      undefinedReferences: "refuse",
      snapshot: createFolioAIEditSnapshot(view.state.doc),
      batch: {
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "op",
            type: "insertAfterBlock",
            blockId: "12345678",
            text: "Item",
            numbering: { kind: "newList", format: "bullet" },
          },
        ],
      },
    });
    expect(result?.skipped).toEqual([]);
    const defined =
      manager.api.getCanonicalDocument()?.package.numbering?.nums ?? [];
    expect(defined).not.toEqual([]);
    const known = getStatedDocumentNumbering(
      manager.getView()?.state ?? panic("No body view"),
    );
    expect(known?.nums.map(({ numId }) => numId)).toEqual(
      defined.map(({ numId }) => numId),
    );
  } finally {
    manager.destroyView();
    bodyHost.remove();
    releaseDom();
  }
});

test("a list a public operation starts in an open header reaches that editor's numbering", () => {
  const releaseDom = acquireHarnessDom();
  const bodyHost = document.createElement("div");
  const storyHost = document.createElement("div");
  document.body.append(bodyHost, storyHost);
  const source: Document = {
    package: { document: { content: [paragraph("12345678", "Body")] } },
  };
  const manager = createHiddenEditorManager({
    getHost: () => bodyHost,
    getDocument: () => source,
    getDocumentContext: () => source,
    getExperimentalSession: () => "canonical",
    getStyles: () => null,
    getExtensionManager: () => undefined,
    getExternalPlugins: () => [],
    getCollaboration: () => undefined,
    getCollaborationModules: () => null,
    getPrecomputedInitialState: () => null,
    getReadOnly: () => false,
    getDocumentIdentity: () => "story-new-list-numbering",
    onTransaction: () => {
      notes.sync();
      parts.sync();
    },
    onCanonicalCompositionEnd: () => {},
    onSelectionChange: () => {},
    onKeyDown: () => false,
    onCopy: () => {},
    onCut: () => {},
    onPaste: () => {},
    onReadOnlyEditAttempt: () => {},
    onEditorViewReady: () => {},
    onEditorViewDestroy: () => {},
    onRemoteSelectionsChange: () => {},
    onSessionRefusal: (message) => panic(message),
  });
  const deps = {
    getHost: () => storyHost,
    getDocument: () => source,
    getCanonicalApi: () => manager.api,
    readCanonicalDocument: () => manager.readCanonicalDocument(),
    getExperimentalSession: () => "canonical" as const,
    getStyles: () => null,
    getTheme: () => null,
  };
  const notes = createNoteEditorManager(deps);
  const parts = createHeaderFooterEditorManager(deps);
  try {
    manager.ensureView();
    expect(
      manager.api.applyCanonicalOperations([
        {
          type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
          sectionIndex: 0,
          story: { kind: "header", rId: "rIdHeader" },
          referenceType: "default",
          content: [paragraph("23456781", "Header")],
        },
      ]),
    ).toBe(true);
    const header = parts.getView("rIdHeader") ?? panic("No header editor");
    header.dispatch(
      header.state.tr.setSelection(TextSelection.create(header.state.doc, 2)),
    );
    const projection =
      manager.api.getCanonicalStoryProjection({
        kind: "header",
        rId: "rIdHeader",
      }) ?? panic("No header projection");
    const result = manager.api.applyCanonicalDocumentOperations({
      undefinedReferences: "refuse",
      story: { type: "header", relationshipId: "rIdHeader" },
      snapshot: createFolioAIEditSnapshot(projection),
      batch: {
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "op",
            type: "insertAfterBlock",
            blockId: "23456781",
            text: "Item",
            numbering: { kind: "newList", format: "bullet" },
          },
        ],
      },
    });
    expect(result?.skipped).toEqual([]);
    parts.sync();
    const defined =
      manager.api.getCanonicalDocument()?.package.numbering?.nums ?? [];
    expect(defined).not.toEqual([]);
    const view = parts.getView("rIdHeader") ?? panic("No header editor");
    expect(view.state.doc.textContent).toBe("HeaderItem");
    expect(
      getStatedDocumentNumbering(view.state)?.nums.map(({ numId }) => numId),
    ).toEqual(defined.map(({ numId }) => numId));
  } finally {
    notes.destroy();
    parts.destroy();
    manager.destroyView();
    bodyHost.remove();
    storyHost.remove();
    releaseDom();
  }
});
