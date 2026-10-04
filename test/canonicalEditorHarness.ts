import { GlobalRegistrator } from "@happy-dom/global-registrator";
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
import { CANONICAL_GAP } from "../packages/core/src/types/canonicalCapabilities";
import type { Document } from "../packages/core/src/types/document";
import { keyboardEventFor, type EditorMode } from "../packages/core/src/__tests__/editorHarness";
import { HARNESS_AUTHOR } from "../packages/core/src/__tests__/editorHarness";
import { singletonManager } from "../packages/core/src/prosemirror/schema";

/** A disposable driver over the production controller, with no PM mutation fallback. */
export const createCanonicalEditorHarness = (source: Document, mode: EditorMode) => {
  if (typeof document === "undefined") GlobalRegistrator.register();
  const host = document.createElement("div");
  document.body.append(host);
  const refusals: HarnessRefusal[] = [];
  const refusalRows: HarnessRefusalRow[] = [];
  let acceptedTransactions = 0;
  let expectedRow: HarnessRefusalRow | undefined;
  const activationRow = canonicalActivationRefusalRow(source);
  const activationSource = cloneDocumentWithParagraphPropertySources(source);
  if (activationRow) validateHarnessRefusalRows([activationRow]);
  const withRefusalRow = <T>(row: HarnessRefusalRow | undefined, action: () => T) => {
    expectedRow = row;
    if (row) {
      validateHarnessRefusalRows([row]);
      refusalRows.push(row);
    }
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
      if (row && row.id !== "missing-command-descriptor" && refusals.length === before.refusals)
        panic(`Canonical refusal row must become strict: ${row.id}`);
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
      expectedRow = undefined;
    }
  };
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
      const row = expectedRow;
      refusals.push({
        gap,
        message,
        expectation: row && gap === row.gap && message === row.message ? "declared" : "unexpected",
        ...(row ? { row: row.id } : {}),
      });
    },
  });
  const dispose = () => {
    manager.destroyView();
    host.remove();
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
    manager.api.getCanonicalDocument() ?? panic("Canonical snapshot unavailable.");
  const execute = (command: Command) => {
    const count = refusals.length;
    const row =
      getCanonicalCommandIntents(command, editorView.state) === undefined
        ? {
            id: "missing-command-descriptor",
            gap: CANONICAL_GAP.dispatch,
            message: "A plugin attempted an unclassified canonical document mutation.",
          }
        : undefined;
    return withRefusalRow(row, () => {
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
  if (ids.length === 0) return { model: session.document, state };
  const commit = session.prepareResolve(state, { revisionIds: ids, resolution: decision }).unwrap();
  const published = publishCanonicalProjection({ state, session, commit }).unwrap();
  return { model: session.document, state: published.state };
};
