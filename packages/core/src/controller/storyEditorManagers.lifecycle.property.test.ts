import { expect, test, setDefaultTimeout } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { panic } from "better-result";
import fc from "fast-check";
import { AllSelection, EditorState, NodeSelection, TextSelection } from "prosemirror-state";
import { DOCUMENT_OP_TYPES, OP_STORIES, type DocumentOp } from "@stll/docx-core/ops";

import { assertExactModel } from "../../../../test/exactModel";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { schema } from "../prosemirror/schema";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import type { Document, Paragraph } from "../types/document";
import {
  createCanonicalSession,
  CanonicalSessionError,
  type CanonicalCommit,
} from "./canonicalSession";
import { createHiddenEditorManager } from "./hiddenEditorManager";
import { createHiddenEditorApi } from "./hiddenEditorApi";
import { createNoteEditorManager, enumerateDocumentNoteStories } from "./noteEditorManager";
import {
  createHeaderFooterEditorManager,
  enumerateDocumentHeaderFooterParts,
} from "./headerFooterEditorManager";

setDefaultTimeout(propertyTestTimeout(120_000));

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const sourceDocument = (): Document => ({
  package: { document: { content: [paragraph("12345678", "Body")] } },
});

const harness = () => {
  const staleHostDocument = sourceDocument();
  const session = createCanonicalSession(staleHostDocument).unwrap();
  let bodyState = EditorState.create({ schema, doc: session.projection.doc });
  let frozen = false;
  let unavailable = false;
  const commit = (prepared: CanonicalCommit) => {
    bodyState = bodyState.apply(prepared.transaction);
    expect(prepared.publish().isOk()).toBe(true);
    expect(bodyState.doc.eq(session.projection.doc)).toBe(true);
  };
  const api = createHiddenEditorApi({
    getView: () => null,
    getDocumentContext: () => staleHostDocument,
    getCanonicalDocument: () => {
      if (frozen) panic("A pending composition snapshot was read.");
      return unavailable ? null : session.document;
    },
    isDestroying: () => false,
    ensureView: () => {},
    isViewRequested: () => true,
    canonicalOperations: {
      updateCanonicalInputLifecycle: () => false,
      applyCanonicalOperations: (ops) => {
        commit(session.prepareOperations(bodyState, ops).unwrap());
        return true;
      },
      getCanonicalStoryProjection: (story) =>
        frozen || unavailable ? null : session.projectStory(story).unwrap().doc,
      getCanonicalStorySelection: () => null,
      replaceCanonicalStoryText: () => false,
      applyCanonicalStoryHistory: () => false,
    },
  });
  const host = document.createElement("div");
  document.body.append(host);
  const deps = {
    getDocument: () => staleHostDocument,
    getCanonicalApi: () => api,
    getExperimentalSession: () => "canonical" as const,
    getHost: () => host,
    getStyles: () => null,
    getTheme: () => null,
  };
  const notes = createNoteEditorManager(deps);
  const parts = createHeaderFooterEditorManager(deps);
  const sync = () => {
    notes.sync();
    parts.sync();
    expect(notes.listStories()).toEqual(enumerateDocumentNoteStories(session.document));
    expect(parts.listSlots()).toEqual(enumerateDocumentHeaderFooterParts(session.document));
    assertExactModel(notes.snapshotDocument(staleHostDocument), session.document);
    assertExactModel(parts.snapshotDocument(staleHostDocument), session.document);
    for (const story of notes.listStories()) {
      const view = notes.activate(story);
      if (!view) panic("An owned canonical note did not mount.");
      expect(
        view.state.doc.eq(
          session.projectStory({ kind: story.kind, id: story.noteId }).unwrap().doc,
        ),
      ).toBe(true);
    }
    for (const story of parts.listSlots()) {
      const view = parts.getView(story.rId);
      if (!view) panic("An owned canonical header/footer did not mount.");
      expect(view.state.doc.eq(session.projectStory(story).unwrap().doc)).toBe(true);
    }
  };
  return {
    session,
    notes,
    parts,
    api,
    sync,
    state: () => bodyState,
    apply: (ops: readonly DocumentOp[]) => {
      api.applyCanonicalOperations(ops);
      sync();
    },
    bodyEdit: (text: string) => {
      commit(session.prepareReplace(bodyState, { from: 1, to: 1, text }).unwrap());
      sync();
    },
    undo: () => {
      commit(session.prepareUndo(bodyState).unwrap());
      sync();
    },
    redo: () => {
      commit(session.prepareRedo(bodyState).unwrap());
      sync();
    },
    freeze: () => {
      frozen = true;
    },
    unfreeze: () => {
      frozen = false;
    },
    refuse: () => {
      unavailable = true;
    },
    resume: () => {
      unavailable = false;
    },
    destroy: () => {
      notes.destroy();
      parts.destroy();
      host.remove();
    },
  };
};

test("generated canonical story lifecycle follows the owner while host props stay stale", async () => {
  GlobalRegistrator.register();
  try {
    await assertProperty(
      fc.asyncProperty(
        fc.record({
          variant: fc.constantFrom("default", "first", "even"),
          id: fc.integer({ min: 2, max: 99 }),
          text: fc.constantFrom("x", "😀", "é", "مرحبا"),
        }),
        async ({ variant, id, text }) => {
          const owner = harness();
          try {
            const snapshots = [owner.session.document];
            const selections = [owner.state().selection.toJSON()];
            const record = () => {
              snapshots.push(owner.session.document);
              selections.push(owner.state().selection.toJSON());
            };
            const at = { story: OP_STORIES.MAIN, blockId: "12345678", offset: 0 } as const;
            owner.apply([
              {
                type: DOCUMENT_OP_TYPES.ADD_NOTE,
                at,
                note: { type: "footnote", id, content: [paragraph("23456789", text)] },
              },
            ]);
            record();
            owner.apply([
              {
                type: DOCUMENT_OP_TYPES.ADD_NOTE,
                at,
                note: { type: "endnote", id, content: [paragraph("34567890", text)] },
              },
            ]);
            record();
            const footer = { kind: "footer", rId: "rIdFooter" } as const;
            owner.apply([
              {
                type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
                sectionIndex: 0,
                story: footer,
                referenceType: variant,
                content: [paragraph("45678901", text)],
              },
            ]);
            record();
            const footView = owner.notes.getView({ kind: "footnote", noteId: id });
            const endView = owner.notes.getView({ kind: "endnote", noteId: id });
            const footerView = owner.parts.getView(footer.rId);
            if (!footView || !endView || !footerView) panic("The lifecycle fixture lost a view.");
            expect(footView === endView).toBe(false);
            owner.freeze();
            expect(() => {
              owner.notes.sync();
              owner.parts.sync();
            }).not.toThrow();
            expect(owner.notes.getView({ kind: "footnote", noteId: id }) === footView).toBe(true);
            expect(owner.parts.getView(footer.rId) === footerView).toBe(true);
            owner.unfreeze();
            owner.refuse();
            expect(() => {
              owner.notes.sync();
              owner.parts.sync();
            }).not.toThrow();
            expect(owner.notes.getView({ kind: "footnote", noteId: id }) === footView).toBe(true);
            expect(owner.parts.getView(footer.rId) === footerView).toBe(true);
            owner.resume();
            owner.sync();
            owner.bodyEdit(text);
            record();
            owner.apply([
              {
                type: DOCUMENT_OP_TYPES.REMOVE_NOTE,
                at: { ...at, offset: text.length },
                story: { kind: "endnote", id },
              },
            ]);
            record();
            expect(endView.isDestroyed).toBe(true);
            expect(owner.notes.getView({ kind: "footnote", noteId: id }) === footView).toBe(true);
            owner.apply([
              {
                type: DOCUMENT_OP_TYPES.REMOVE_NOTE,
                at: { ...at, offset: text.length },
                story: { kind: "footnote", id },
              },
            ]);
            record();
            expect(footView.isDestroyed).toBe(true);
            expect(owner.notes.getActive()).toBeNull();
            owner.apply([
              {
                type: DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER,
                sectionIndex: 0,
                story: footer,
                referenceType: variant,
              },
            ]);
            record();
            expect(footerView.isDestroyed).toBe(true);
            for (let index = snapshots.length - 2; index >= 0; index--) {
              owner.undo();
              assertExactModel(owner.session.document, snapshots.at(index));
              expect(owner.state().selection.toJSON()).toEqual(selections.at(index));
            }
            for (let index = 1; index < snapshots.length; index++) {
              owner.redo();
              assertExactModel(owner.session.document, snapshots.at(index));
              expect(owner.state().selection.toJSON()).toEqual(selections.at(index));
            }
            owner.refuse();
            expect(() => {
              owner.notes.sync();
              owner.parts.sync();
            }).not.toThrow();
            expect(owner.notes.listStories()).toEqual([]);
            for (const manager of [owner.notes, owner.parts]) {
              const snapshot = () => manager.snapshotDocument(sourceDocument());
              expect(snapshot).toThrow(CanonicalSessionError);
              let error: unknown;
              try {
                snapshot();
              } catch (refusal) {
                error = refusal;
              }
              if (!(error instanceof CanonicalSessionError))
                panic("The unavailable snapshot lost its typed refusal.");
              expect(error.gap).toBe(CANONICAL_GAP.authorityRouting);
            }
          } finally {
            owner.destroy();
          }
        },
      ),
      {
        numRuns: 12,
        id: "generated canonical story lifecycle follows the owner while host props stay stale",
      },
    );
  } finally {
    GlobalRegistrator.unregister();
  }
});

for (const target of [
  { kind: "footer", variant: "default" },
  { kind: "footer", variant: "first" },
  { kind: "footer", variant: "even" },
  { kind: "footnote" },
  { kind: "endnote" },
] as const) {
  for (const editText of [false, true]) {
    for (const selectionType of ["text", "all", "node"] as const) {
      test(`native ${editText ? "typed " : ""}${target.kind}${target.kind === "footer" ? ` ${target.variant}` : ""} ${selectionType === "node" ? "creation refuses without mutation" : "undo retires its active newly-created story"} with ${selectionType} selection`, () => {
        GlobalRegistrator.register();
        const bodyHost = document.createElement("div");
        const storyHost = document.createElement("div");
        document.body.append(bodyHost, storyHost);
        const source = sourceDocument();
        const refusals: string[] = [];
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
          getDocumentIdentity: () => "native-story-history",
          onTransaction: () => {
            notes.sync();
            parts.sync();
          },
          onSelectionChange: () => {},
          onKeyDown: () => false,
          onCopy: () => {},
          onCut: () => {},
          onPaste: () => {},
          onReadOnlyEditAttempt: () => {},
          onEditorViewReady: () => {},
          onEditorViewDestroy: () => {},
          onRemoteSelectionsChange: () => {},
          onSessionRefusal: (message) => {
            refusals.push(message);
          },
        });
        const deps = {
          getHost: () => storyHost,
          getDocument: () => source,
          getCanonicalApi: () => manager.api,
          getExperimentalSession: () => "canonical" as const,
          getStyles: () => null,
          getTheme: () => null,
        };
        const notes = createNoteEditorManager(deps);
        const parts = createHeaderFooterEditorManager(deps);
        try {
          manager.ensureView();
          const bodyView = manager.getView() ?? panic("Expected native body view.");
          const selections = {
            text: () => TextSelection.create(bodyView.state.doc, 2, 4),
            all: () => new AllSelection(bodyView.state.doc),
            node: () => NodeSelection.create(bodyView.state.doc, 0),
          };
          bodyView.dispatch(bodyView.state.tr.setSelection(selections[selectionType]()));
          const original = manager.api.getCanonicalDocument();
          const originalSelection = manager.getView()?.state.selection.toJSON();
          const story =
            target.kind === "footer"
              ? ({ kind: "footer", rId: "rIdNewFooter" } as const)
              : ({ kind: target.kind, id: 9 } as const);
          const operation =
            target.kind === "footer"
              ? ({
                  type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
                  sectionIndex: 0,
                  story: { kind: "footer", rId: "rIdNewFooter" },
                  referenceType: target.variant,
                  content: [paragraph("23456789", "Story")],
                } as const satisfies DocumentOp)
              : ({
                  type: DOCUMENT_OP_TYPES.ADD_NOTE,
                  at: { story: OP_STORIES.MAIN, blockId: "12345678", offset: 0 },
                  note: { type: target.kind, id: 9, content: [paragraph("23456789", "Story")] },
                } as const satisfies DocumentOp);
          if (selectionType === "node") {
            expect(manager.api.applyCanonicalOperations([operation])).toBe(false);
            assertExactModel(manager.api.getCanonicalDocument(), original);
            expect(manager.getView()?.state.selection.toJSON()).toEqual(originalSelection);
            expect(notes.listStories()).toEqual([]);
            expect(parts.listSlots()).toEqual([]);
            expect(refusals).toHaveLength(1);
            expect(refusals.at(0)).toContain("selection");
            expect(manager.api.undo()).toBe(false);
            return;
          }
          expect(manager.api.applyCanonicalOperations([operation])).toBe(true);
          const created = manager.api.getCanonicalDocument();
          const storyView =
            target.kind === "footer"
              ? parts.getView("rIdNewFooter")
              : notes.activate({ kind: target.kind, noteId: 9 });
          if (!storyView) panic("The native history fixture did not mount its story.");
          if (editText) {
            const position = storyView.state.doc.content.size - 1;
            const accepted = manager.api.replaceCanonicalStoryText({
              view: storyView,
              story,
              intent: { from: position, to: position, text: "X", semantic: "replacement" },
            });
            expect(refusals).toEqual([]);
            expect(accepted).toBe(true);
          }
          const edited = manager.api.getCanonicalDocument();
          for (const expected of editText ? [created, original] : [original]) {
            const event = new KeyboardEvent("keydown", {
              key: "z",
              ctrlKey: true,
              metaKey: true,
              bubbles: true,
              cancelable: true,
            });
            storyView.dom.dispatchEvent(event);
            expect(event.defaultPrevented).toBe(true);
            assertExactModel(manager.api.getCanonicalDocument(), expected);
          }
          expect(storyView.isDestroyed).toBe(true);
          if (target.kind === "footer") expect(parts.getView("rIdNewFooter")).toBeNull();
          else {
            expect(notes.getView({ kind: target.kind, noteId: 9 })).toBeNull();
            expect(notes.getActive()).toBeNull();
          }
          expect(manager.getView()?.state.selection.toJSON()).toEqual(originalSelection);
          expect(manager.api.redo()).toBe(true);
          assertExactModel(manager.api.getCanonicalDocument(), created);
          const restoredView =
            target.kind === "footer"
              ? parts.getView("rIdNewFooter")
              : notes.activate({ kind: target.kind, noteId: 9 });
          if (!restoredView) panic("Redo did not remount its restored story.");
          if (editText)
            expect(
              manager.api.applyCanonicalStoryHistory({
                view: restoredView,
                story,
                direction: "redo",
              }),
            ).toBe(true);
          assertExactModel(manager.api.getCanonicalDocument(), edited);
          expect(refusals).toEqual([]);
        } finally {
          notes.destroy();
          parts.destroy();
          manager.destroyView();
          bodyHost.remove();
          storyHost.remove();
          GlobalRegistrator.unregister();
        }
      });
    }
  }
}
