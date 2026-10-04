import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { panic } from "better-result";
import type { Slice } from "prosemirror-model";
import type { Command, Transaction } from "prosemirror-state";

import { storyRevisionIds } from "./reviewProjection";
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
import type { CanonicalGap } from "../packages/core/src/types/canonicalCapabilities";
import type { Document } from "../packages/core/src/types/document";
import { keyboardEventFor, type EditorMode } from "../packages/core/src/__tests__/editorHarness";
import { HARNESS_AUTHOR } from "../packages/core/src/__tests__/editorHarness";
import { singletonManager } from "../packages/core/src/prosemirror/schema";

export type HarnessRefusal = {
  gap: CanonicalGap;
  message: string;
  expectation: "missing-command-descriptor" | "unexpected";
};

/** A disposable driver over the production controller, with no PM mutation fallback. */
export const createCanonicalEditorHarness = (source: Document, mode: EditorMode) => {
  if (typeof document === "undefined") GlobalRegistrator.register();
  const host = document.createElement("div");
  document.body.append(host);
  const refusals: HarnessRefusal[] = [];
  let expectedGap: CanonicalGap | undefined;
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
    onTransaction: () => {},
    onSelectionChange: () => {},
    onKeyDown: () => false,
    onReadOnlyEditAttempt: () => {},
    onEditorViewReady: () => {},
    onEditorViewDestroy: () => {},
    onRemoteSelectionsChange: () => {},
    onSessionRefusal: (message, gap) => {
      refusals.push({
        gap,
        message,
        expectation: gap === expectedGap ? "missing-command-descriptor" : "unexpected",
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
    dispose();
    const refusal = refusals.at(0);
    if (refusal) throw new CanonicalSessionRefusalError(refusal);
    panic("The canonical conformance controller did not create its view.");
  }
  const snapshot = () =>
    manager.api.getCanonicalDocument() ?? panic("Canonical snapshot unavailable.");
  const execute = (command: Command) => {
    const count = refusals.length;
    expectedGap =
      getCanonicalCommandIntents(command, editorView.state) === undefined
        ? CANONICAL_GAP.dispatch
        : undefined;
    try {
      const applied = executeEditorCommand(editorView, command);
      return refusals.length === count && applied;
    } finally {
      expectedGap = undefined;
    }
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
    snapshot,
    history: manager.api,
    dispose,
    resolve,
    pressKey: (binding: string) => {
      const fields = keyboardEventFor(binding);
      const event = new KeyboardEvent("keydown", { ...fields, bubbles: true, cancelable: true });
      editorView.dom.dispatchEvent(event);
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
      editorView.someProp("handlePaste", (handler) => handler(editorView, event, slice));
    },
    cut: () => {
      const event = new ClipboardEvent("cut", { bubbles: true, cancelable: true });
      editorView.dom.dispatchEvent(event);
      return event.defaultPrevented;
    },
  };
};

export type CanonicalEditorHarness = ReturnType<typeof createCanonicalEditorHarness>;

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
