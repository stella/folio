import { acquireHarnessDom } from "./harnessDom";
import { panic, Result } from "better-result";
import type { Slice } from "prosemirror-model";
import type { Command, Transaction } from "prosemirror-state";

import {
  canonicalActivationRefusalRow,
  validateHarnessRefusalRows,
  type HarnessRefusal,
  type HarnessRefusalRow,
} from "./canonical-refusal-rows";
import { storyRevisionIds } from "./reviewProjection";
import { assertExactModel } from "./exactModel";
import { cloneDocumentWithParagraphPropertySources } from "../packages/core/src/docx/paragraphPropertySource";
import {
  createHiddenEditorManager,
  CanonicalSessionRefusalError,
} from "../packages/core/src/controller/hiddenEditorManager";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../packages/core/src/controller/canonicalSession";
import { EditorState as PMEditorState } from "prosemirror-state";
import { executeEditorCommand } from "../packages/core/src/prosemirror/executeEditorCommand";
import { getCanonicalCommandIntents } from "../packages/core/src/prosemirror/canonicalCommands";
import { canonicalClipboardStoryPartRefusal } from "../packages/core/src/controller/canonicalClipboard";
import { proseDocToBlocks } from "../packages/core/src/prosemirror/conversion/fromProseDoc";
import { CANONICAL_GAP } from "../packages/core/src/types/canonicalCapabilities";
import { serializeCanonicalSave } from "../packages/core/src/docx/canonicalSave";
import type { CanonicalSaveSnapshot } from "../packages/core/src/types/canonicalSave";
import type { Document } from "../packages/core/src/types/document";
import { keyboardEventFor, type EditorMode } from "../packages/core/src/__tests__/editorHarness";
import { HARNESS_AUTHOR } from "../packages/core/src/__tests__/editorHarness";
import { singletonManager } from "../packages/core/src/prosemirror/schema";

const canonicalSaveSnapshots = new WeakMap<Document, CanonicalSaveSnapshot>();

const rememberCanonicalSave = (snapshot: CanonicalSaveSnapshot): Document => {
  canonicalSaveSnapshots.set(snapshot.document, snapshot);
  return snapshot.document;
};

/** Serialize the committed snapshot through the same owner as the adapters. */
export const saveCanonicalHarnessDocument = async (source: Document) => {
  const snapshot =
    canonicalSaveSnapshots.get(source) ??
    createCanonicalSession(source).unwrap().captureSaveSnapshot();
  const saved = await serializeCanonicalSave({ snapshot, options: { mode: "full" } });
  return { model: saved.document, bytes: new Uint8Array(saved.buffer) };
};

/** A disposable driver over the production controller, with no PM mutation fallback. */
export const createCanonicalEditorHarness = (source: Document, mode: EditorMode) => {
  const releaseDom = acquireHarnessDom();
  const host = document.createElement("div");
  document.body.append(host);
  const refusals: HarnessRefusal[] = [];
  const refusalRows: HarnessRefusalRow[] = [];
  let acceptedTransactions = 0;
  let expectedRows: readonly HarnessRefusalRow[] = [];
  const activationRow = canonicalActivationRefusalRow(source);
  const activationSource = cloneDocumentWithParagraphPropertySources(source);
  if (activationRow) validateHarnessRefusalRows([activationRow]);
  const withRefusalRows = <T>(rows: readonly HarnessRefusalRow[], action: () => T) => {
    expectedRows = rows;
    validateHarnessRefusalRows(rows);
    refusalRows.push(...rows);
    const before = {
      document: manager.api.getCanonicalDocument(),
      state: editorView.state,
      canUndo: manager.api.canUndo(),
      canRedo: manager.api.canRedo(),
      refusals: refusals.length,
      acceptedTransactions,
    };
    try {
      const result = action();
      if (
        rows.some(({ id }) => id !== "missing-command-descriptor") &&
        refusals.length === before.refusals
      )
        panic(`Canonical refusal rows must become strict: ${rows.map(({ id }) => id).join(", ")}`);
      if (refusals.length > before.refusals) {
        const after = manager.api.getCanonicalDocument();
        if (!after || !before.document) panic("Refusal lost canonical authority");
        assertExactModel(after, before.document);
        if (
          !editorView.state.doc.eq(before.state.doc) ||
          !editorView.state.selection.eq(before.state.selection) ||
          manager.api.canUndo() !== before.canUndo ||
          manager.api.canRedo() !== before.canRedo ||
          acceptedTransactions !== before.acceptedTransactions
        )
          panic("Canonical refusal changed projection, selection or journal availability");
      }
      return result;
    } finally {
      expectedRows = [];
    }
  };
  const withRefusalRow = <T>(row: HarnessRefusalRow | undefined, action: () => T) =>
    withRefusalRows(row ? [row] : [], action);
  const manager = createHiddenEditorManager({
    getHost: () => host,
    getDocument: () => source,
    getDocumentContext: () => source,
    getExperimentalSession: () => "canonical",
    getEditingMode: () => mode,
    getSuggestionAuthor: () => HARNESS_AUTHOR,
    getStyles: () => source.package.styles,
    getExtensionManager: () => singletonManager,
    getExternalPlugins: () => [],
    getCollaboration: () => undefined,
    getCollaborationModules: () => null,
    getPrecomputedInitialState: () => null,
    getReadOnly: () => false,
    getDocumentIdentity: () => "conformance",
    onTransaction: () => {
      acceptedTransactions += 1;
    },
    onSelectionChange: () => {},
    onKeyDown: () => false,
    onReadOnlyEditAttempt: () => {},
    onEditorViewReady: () => {},
    onEditorViewDestroy: () => {},
    onRemoteSelectionsChange: () => {},
    onSessionRefusal: (message, gap) => {
      const row = expectedRows.find(
        (candidate) => candidate.gap === gap && candidate.message === message,
      );
      refusals.push({
        gap,
        message,
        expectation: row && gap === row.gap && message === row.message ? "declared" : "unexpected",
        ...(row ? { row: row.id } : {}),
      });
    },
  });
  let disposed = false;
  const dispose = () => {
    if (disposed) return;
    disposed = true;
    try {
      manager.destroyView();
    } finally {
      host.remove();
      releaseDom();
    }
  };
  manager.ensureView();
  const editorView = manager.getView();
  if (!editorView) {
    const refusal = refusals.at(0);
    if (activationRow) {
      assertExactModel(source, activationSource);
      if (
        manager.api.getCanonicalDocument() !== null ||
        manager.api.canUndo() ||
        manager.api.canRedo() ||
        refusal?.gap !== activationRow.gap ||
        refusal.message !== activationRow.message
      ) {
        dispose();
        panic("Canonical activation refusal changed authority or lost its ledger contract");
      }
    }
    dispose();
    if (refusal) throw new CanonicalSessionRefusalError(refusal);
    panic("The canonical conformance controller did not create its view.");
  }
  if (activationRow) {
    dispose();
    panic("Canonical activation refusal row must become strict");
  }
  const snapshot = () =>
    rememberCanonicalSave(
      manager.api.captureCanonicalSave() ?? panic("Canonical save snapshot unavailable."),
    );
  const execute = (command: Command) => {
    const count = refusals.length;
    const intents = getCanonicalCommandIntents(command, editorView.state);
    const rows =
      intents === undefined
        ? [
            "A plugin attempted an unclassified canonical document mutation.",
            "Unclassified native text is unavailable in this session.",
          ].map((message) => ({
            id: "missing-command-descriptor",
            gap: CANONICAL_GAP.dispatch,
            message,
          }))
        : [];
    if (
      mode === "suggesting" &&
      intents?.some(
        ({ type }) =>
          type === "setHyperlink" || type === "removeHyperlink" || type === "insertHyperlink",
      )
    )
      rows.push({
        id: "tracked-hyperlink-resolution",
        gap: CANONICAL_GAP.trackedHyperlinkResolution,
        message: "Hyperlink suggestions require serializable wrapper review provenance.",
      });
    return withRefusalRows(rows, () => {
      const applied = executeEditorCommand(editorView, command);
      return refusals.length === count && applied;
    });
  };
  const resolve = (decision: "accept" | "reject") => {
    const ids = storyRevisionIds(snapshot());
    return ids.length > 0 && manager.api.resolveCanonicalRevisions(ids, decision);
  };
  return {
    authority: "canonical" as const,
    commandManager: singletonManager,
    get state() {
      return editorView.state;
    },
    dispatch: (transaction: Transaction) => editorView.dispatch(transaction),
    execute,
    refusals,
    refusalRows,
    snapshot,
    history: manager.api,
    dispose,
    resolve,
    pressKey: (binding: string) => {
      const fields = keyboardEventFor(binding);
      const event = new KeyboardEvent("keydown", { ...fields, bubbles: true, cancelable: true });
      const row =
        (fields.key === "Backspace" || fields.key === "Delete") &&
        (fields.ctrlKey || fields.metaKey || fields.altKey)
          ? {
              id: "modified-deletion",
              gap: CANONICAL_GAP.dispatch,
              message: "Only plain character deletion is available in this session.",
            }
          : undefined;
      withRefusalRow(row, () => editorView.dom.dispatchEvent(event));
      return event.defaultPrevented;
    },
    typeText: (text: string) => {
      for (const character of text)
        editorView.dom.dispatchEvent(
          new InputEvent("beforeinput", {
            inputType: "insertText",
            data: character,
            bubbles: true,
            cancelable: true,
          }),
        );
    },
    paste: (slice: Slice) => {
      const event = new ClipboardEvent("paste", { bubbles: true, cancelable: true });
      let paragraphsOnly = true;
      let inlineOnly = true;
      slice.content.forEach((node) => {
        paragraphsOnly &&= node.type.name === "paragraph";
        inlineOnly &&= node.isInline;
      });
      let row: HarnessRefusalRow | undefined;
      if (slice.openStart > 1 || slice.openEnd > 1)
        row = {
          id: "clipboard-nested-blocks",
          gap: CANONICAL_GAP.dispatch,
          message: "Clipboard tables and nested block containers require canonical table editing.",
        };
      else if (!paragraphsOnly && !inlineOnly)
        row = {
          id: "clipboard-table",
          gap: CANONICAL_GAP.dispatch,
          message: "Clipboard tables and embedded blocks require canonical table editing.",
        };
      if (row === undefined) {
        const paragraphType = editorView.state.schema.nodes["paragraph"];
        if (paragraphType === undefined) panic("Canonical clipboard schema lost paragraphs");
        const content = inlineOnly ? paragraphType.create(null, slice.content) : slice.content;
        const blocks = proseDocToBlocks(
          editorView.state.schema.topNodeType.create(null, content),
          [],
        );
        for (const block of blocks) {
          if (block.type !== "paragraph") continue;
          const refusal = canonicalClipboardStoryPartRefusal(block);
          if (refusal === undefined) continue;
          row = { id: "clipboard-story-parts", gap: refusal.gap, message: refusal.message };
          break;
        }
      }
      withRefusalRow(row, () =>
        editorView.someProp("handlePaste", (handler) => handler(editorView, event, slice)),
      );
    },
    cut: () => {
      const event = new ClipboardEvent("cut", {
        bubbles: true,
        cancelable: true,
        clipboardData: new DataTransfer(),
      });
      editorView.dom.dispatchEvent(event);
      return event.defaultPrevented;
    },
  };
};

export type CanonicalEditorHarness = ReturnType<typeof createCanonicalEditorHarness>;

/** Unsupported activation is a counted precondition; every other initialization failure is strict. */
export const createCanonicalHarnessCase = (source: Document, mode: EditorMode) => {
  const created = Result.try({
    try: () => createCanonicalEditorHarness(source, mode),
    catch: (error) => error,
  });
  if (created.isOk()) return { type: "ready", driver: created.value } as const;
  const row = canonicalActivationRefusalRow(source);
  if (
    row &&
    created.error instanceof CanonicalSessionRefusalError &&
    created.error.gap === row.gap &&
    created.error.message === row.message
  )
    return {
      type: "activationRefused",
      refusal: { gap: row.gap, message: row.message, expectation: "declared", row: row.id },
    } as const;
  throw created.error;
};

/** Review observations use a fresh canonical session, leaving the case's journal untouched. */
export const resolveCanonicalHarnessDocument = (
  source: Document,
  decision: "accept" | "reject",
) => {
  const session = createCanonicalSession(source).unwrap();
  const state = PMEditorState.create({ doc: session.projection.doc });
  const ids = storyRevisionIds(session.document);
  if (ids.length === 0)
    return { model: rememberCanonicalSave(session.captureSaveSnapshot()), state };
  const commit = session.prepareResolve(state, { revisionIds: ids, resolution: decision }).unwrap();
  const published = publishCanonicalProjection({ state, session, commit }).unwrap();
  return { model: rememberCanonicalSave(session.captureSaveSnapshot()), state: published.state };
};
