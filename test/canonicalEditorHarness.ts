import { panic, Result } from "better-result";
import type { Node as PMNode, Slice } from "prosemirror-model";
import type { Command, Transaction } from "prosemirror-state";
import { EditorState as PMEditorState } from "prosemirror-state";

import {
  type EditorMode,
  HARNESS_AUTHOR,
  keyboardEventFor,
  parseShapeDocument,
} from "../packages/core/src/__tests__/editorHarness";
import { FolioDocxReviewer } from "../packages/core/src/ai-edits/headless";
import { getTrackedChangesFromDoc } from "../packages/core/src/ai-edits/read";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../packages/core/src/controller/canonicalSession";
import {
  CanonicalSessionRefusalError,
  createHiddenEditorManager,
} from "../packages/core/src/controller/hiddenEditorManager";
import { serializeCanonicalSave } from "../packages/core/src/docx/canonicalSave";
import { cloneDocumentWithParagraphPropertySources } from "../packages/core/src/docx/paragraphPropertySource";
import { executeEditorCommand } from "../packages/core/src/prosemirror/executeEditorCommand";
import { expectRangeAnchorAttrs } from "../packages/core/src/prosemirror/rangeAnchorAttrs";
import { singletonManager } from "../packages/core/src/prosemirror/schema";
import { CANONICAL_GAP } from "../packages/core/src/types/canonicalCapabilities";
import type { CanonicalSaveSnapshot } from "../packages/core/src/types/canonicalSave";
import type { Document } from "../packages/core/src/types/document";
import {
  canonicalActivationRefusalRow,
  type HarnessRefusal,
  type HarnessRefusalRow,
  validateHarnessRefusalRows,
} from "./canonical-refusal-rows";
import { assertExactModel } from "./exactModel";
import { acquireHarnessDom } from "./harnessDom";
import { storyRevisionIds } from "./reviewProjection";

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
  const saved = await serializeCanonicalSave({
    snapshot,
    options: { mode: "full" },
  });
  return { model: saved.document, bytes: new Uint8Array(saved.buffer) };
};

/** A disposable driver over the production controller, with no PM mutation fallback. */
export const createCanonicalEditorHarness = (
  source: Document,
  mode: EditorMode,
) => {
  let caseRefusalRows: readonly HarnessRefusalRow[] = [];
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
  const withRefusalRows = <T>(
    rows: readonly HarnessRefusalRow[],
    action: () => T,
  ) => {
    expectedRows = rows;
    validateHarnessRefusalRows(rows);
    refusalRows.push(...rows);
    const before = {
      document: manager.api.getCanonicalDocument(),
      save: manager.api.captureCanonicalSave(),
      state: editorView.state,
      canUndo: manager.api.canUndo(),
      canRedo: manager.api.canRedo(),
      refusals: refusals.length,
      acceptedTransactions,
    };
    try {
      const result = action();
      if (
        rows.some(({ id }) => id !== CANONICAL_GAP.commands) &&
        refusals.length === before.refusals
      )
        panic(
          `Canonical refusal rows must become strict: ${rows.map(({ id }) => id).join(", ")}`,
        );
      if (rows.length > 0 && refusals.length === before.refusals) {
        // Explicit descriptor-gap cases may be inapplicable, but gaining
        // mutation support retires their refusal contract.
        if (!editorView.state.doc.eq(before.state.doc))
          panic(
            `Canonical refusal rows must become strict: ${rows.map(({ id }) => id).join(", ")}`,
          );
        const after = manager.api.getCanonicalDocument();
        if (!after || !before.document)
          panic("Refusal contract lost canonical authority");
        assertExactModel(after, before.document);
      }
      if (refusals.length > before.refusals) {
        const after = manager.api.getCanonicalDocument();
        if (!after || !before.document)
          panic("Refusal lost canonical authority");
        assertExactModel(after, before.document);
        const saved = manager.api.captureCanonicalSave();
        if (
          saved?.version !== before.save?.version ||
          saved?.structure !== before.save?.structure ||
          JSON.stringify(saved?.changedBlockIds) !==
            JSON.stringify(before.save?.changedBlockIds) ||
          !editorView.state.doc.eq(before.state.doc) ||
          !editorView.state.selection.eq(before.state.selection) ||
          manager.api.canUndo() !== before.canUndo ||
          manager.api.canRedo() !== before.canRedo ||
          acceptedTransactions !== before.acceptedTransactions
        )
          panic(
            "Canonical refusal changed projection, selection or journal availability",
          );
      }
      return result;
    } finally {
      expectedRows = [];
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
    onSessionRefusal: (message, ground) => {
      // A rule refusal never matches a ledger row, so it reads as unexpected.
      const { gap } = ground;
      const row = expectedRows.find(
        (candidate) => candidate.gap === gap && candidate.message === message,
      );
      refusals.push({
        gap,
        ...(ground.rule === undefined ? {} : { rule: ground.rule }),
        message,
        expectation:
          row && gap === row.gap && message === row.message
            ? "declared"
            : "unexpected",
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
        panic(
          "Canonical activation refusal changed authority or lost its ledger contract",
        );
      }
    }
    dispose();
    if (refusal)
      throw new CanonicalSessionRefusalError({
        ground:
          refusal.rule === undefined
            ? { gap: refusal.gap ?? CANONICAL_GAP.dispatch }
            : { rule: refusal.rule },
        message: refusal.message,
      });
    panic("The canonical conformance controller did not create its view.");
  }
  if (activationRow) {
    dispose();
    panic("Canonical activation refusal row must become strict");
  }
  const snapshot = () =>
    rememberCanonicalSave(
      manager.api.captureCanonicalSave() ??
        panic("Canonical save snapshot unavailable."),
    );
  const execute = (command: Command) => {
    const count = refusals.length;
    return withRefusalRows(caseRefusalRows, () => {
      const applied = executeEditorCommand(editorView, command);
      return refusals.length === count && applied;
    });
  };
  /** Resolve what the review list shows, by the ids it lists, as a host does. */
  const resolve = (decision: ReviewDecision) => {
    const before = snapshot();
    const listed = listedRevisionIds(editorView.state.doc);
    const resolved =
      listed.length > 0 &&
      manager.api.resolveCanonicalRevisions(listed, decision);
    const stranded = strandedRevisionIds(before, snapshot(), decision);
    if (stranded.length > 0)
      panic(strandedRevisionText({ listed, stranded, decision }));
    return resolved;
  };
  return {
    authority: "canonical" as const,
    commandManager: singletonManager,
    expectRefusals: (rows: readonly HarnessRefusalRow[]) => {
      validateHarnessRefusalRows(rows);
      caseRefusalRows = rows;
    },
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
      const event = new KeyboardEvent("keydown", {
        ...fields,
        bubbles: true,
        cancelable: true,
      });
      const row =
        (fields.key === "Backspace" || fields.key === "Delete") &&
        (fields.ctrlKey || fields.metaKey || fields.altKey)
          ? {
              id: "modified-deletion",
              gap: CANONICAL_GAP.dispatch,
              message:
                "Only plain character deletion is available in this session.",
            }
          : undefined;
      const rows = row ? [row] : caseRefusalRows;
      withRefusalRows(rows, () => editorView.dom.dispatchEvent(event));
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
      const event = new ClipboardEvent("paste", {
        bubbles: true,
        cancelable: true,
      });
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
          message:
            "Clipboard tables and nested block containers require canonical table editing.",
        };
      else if (!paragraphsOnly && !inlineOnly)
        row = {
          id: "clipboard-table",
          gap: CANONICAL_GAP.dispatch,
          message:
            "Clipboard tables and embedded blocks require canonical table editing.",
        };
      if (row === undefined) {
        // Read clipboard carriers independently of the production preflight.
        // A numbered paragraph alone never imports another story's payload.
        let importsStoryParts = false;
        slice.content.descendants((node) => {
          if (
            (node.type.name === "rangeAnchor" &&
              expectRangeAnchorAttrs(node).start.type ===
                "commentRangeStart") ||
            node.type.name === "commentReference" ||
            node.marks.some(
              ({ type }) =>
                type.name === "comment" || type.name === "footnoteRef",
            )
          )
            importsStoryParts = true;
        });
        if (importsStoryParts)
          row = {
            id: "clipboard-story-parts",
            gap: CANONICAL_GAP.dispatch,
            message:
              "Clipboard comments and note references require importing their source story parts.",
          };
      }
      const rows = row ? [row] : caseRefusalRows;
      withRefusalRows(rows, () =>
        editorView.someProp("handlePaste", (handler) =>
          handler(editorView, event, slice),
        ),
      );
    },
    /**
     * Paste clipboard text through the view's own paste event, as a browser
     * does: the editor parses the text against the caret, then pastes it.
     */
    pasteText: (text: string) => {
      const clipboardData = new DataTransfer();
      clipboardData.setData("text/plain", text);
      const event = new ClipboardEvent("paste", {
        bubbles: true,
        cancelable: true,
        clipboardData,
      });
      withRefusalRows(caseRefusalRows, () =>
        editorView.dom.dispatchEvent(event),
      );
      return event.defaultPrevented;
    },
    cut: () => {
      const event = new ClipboardEvent("cut", {
        bubbles: true,
        cancelable: true,
        clipboardData: new DataTransfer(),
      });
      withRefusalRows(caseRefusalRows, () =>
        editorView.dom.dispatchEvent(event),
      );
      return event.defaultPrevented;
    },
  };
};

export type CanonicalEditorHarness = ReturnType<
  typeof createCanonicalEditorHarness
>;

/** Unsupported activation is a counted precondition; every other initialization failure is strict. */
export const createCanonicalHarnessCase = (
  source: Document,
  mode: EditorMode,
) => {
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
      refusal: {
        gap: row.gap,
        message: row.message,
        expectation: "declared",
        row: row.id,
      },
    } as const;
  throw created.error;
};

type ReviewDecision = "accept" | "reject";

/** The ids a host's review list shows: one per listed change, as the live editor lists them. */
const listedRevisionIds = (doc: PMNode) =>
  getTrackedChangesFromDoc(doc).map(({ id }) => id);

/** Resolve `ids` in a fresh canonical session, leaving any case journal untouched. */
const resolveInFreshSession = (
  source: Document,
  ids: (doc: PMNode) => readonly number[],
  decision: ReviewDecision,
) => {
  const session = createCanonicalSession(source).unwrap();
  const state = PMEditorState.create({ doc: session.projection.doc });
  const revisionIds = ids(state.doc);
  if (revisionIds.length === 0)
    return Result.ok({
      model: rememberCanonicalSave(session.captureSaveSnapshot()),
      state,
      revisionIds,
    });
  return session
    .prepareResolve(state, { revisionIds, resolution: decision })
    .andThen((commit) => publishCanonicalProjection({ state, session, commit }))
    .map((published) => ({
      model: rememberCanonicalSave(session.captureSaveSnapshot()),
      state: published.state,
      revisionIds,
    }));
};

/**
 * Revisions a resolution by the listed ids left pending that resolving every
 * revision id in the story resolves: a host has no other way to reach them.
 * A change a lock holds stays pending either way and is not stranded.
 */
const strandedRevisionIds = (
  source: Document,
  resolved: Document,
  decision: ReviewDecision,
): number[] => {
  const left = storyRevisionIds(resolved);
  if (left.length === 0) return [];
  const everything = resolveInFreshSession(
    source,
    () => storyRevisionIds(source),
    decision,
  );
  const held = new Set(
    storyRevisionIds(everything.isOk() ? everything.value.model : source),
  );
  return left.filter((id) => !held.has(id));
};

type StrandedRevisionTextOptions = {
  listed: readonly number[];
  stranded: readonly number[];
  decision: ReviewDecision;
};
const strandedRevisionText = ({
  listed,
  stranded,
  decision,
}: StrandedRevisionTextOptions) =>
  `${decision} of the listed changes ${JSON.stringify(listed)} left revisions ${JSON.stringify(stranded)} pending`;

type ListedResolution = {
  model: Document;
  state: PMEditorState;
  /** The ids the host listed and resolved. */
  listed: readonly number[];
  /** Revisions the listed resolution left that resolving every id resolves. */
  stranded: readonly number[];
};

/** Resolve the ids `list` shows over `source`, and what that strands. */
const resolveListedChanges = (
  source: Document,
  list: (doc: PMNode) => readonly number[],
  decision: ReviewDecision,
): ListedResolution => {
  const { model, state, revisionIds } = resolveInFreshSession(
    source,
    list,
    decision,
  ).unwrap();
  return {
    model,
    state,
    listed: revisionIds,
    stranded: strandedRevisionIds(source, model, decision),
  };
};

/**
 * Review observations of the session's own model resolve the changes the
 * live editor lists, by the ids it lists, in a fresh canonical session.
 */
export const resolveCanonicalHarnessDocument = (
  source: Document,
  decision: ReviewDecision,
): ListedResolution =>
  resolveListedChanges(source, listedRevisionIds, decision);

/**
 * Review observations of the saved file: save it, reopen it with the
 * headless reviewer, and resolve the ids its change list shows in a canonical
 * session over the reopened package, as a host reviewing the file does.
 */
export const resolveReopenedCanonicalHarnessDocument = async (
  source: Document,
  decision: ReviewDecision,
): Promise<ListedResolution> => {
  // Nothing to list: the reopened file resolves to itself, so skip the reopen.
  if (storyRevisionIds(source).length === 0)
    return resolveListedChanges(source, () => [], decision);
  const { bytes } = await saveCanonicalHarnessDocument(source);
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer);
  const listed = reviewer.getChanges().map(({ id }) => id);
  return resolveListedChanges(
    await parseShapeDocument(bytes),
    () => listed,
    decision,
  );
};
