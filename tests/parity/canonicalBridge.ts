import { TextSelection } from "prosemirror-state";

import type { Document } from "@stll/folio-core";
import { createFolioAIEditSnapshot } from "@stll/folio-core/ai-edits/snapshot";
import type { FolioAIEditSnapshot } from "@stll/folio-core/ai-edits/types";
import type { OpStory } from "@stll/folio-core/controller/canonicalOperations";
import type { FolioEditor } from "@stll/folio-core/controller/folioEditor";
import type { FolioDocumentOperationUndoHandle } from "@stll/folio-core/document-operations";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationResult,
} from "@stll/folio-core/document-operations";
import { applyCellSelection } from "@stll/folio-core/prosemirror/cellDragSelection";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import { singletonManager } from "@stll/folio-core/prosemirror/schema";

import {
  getDocumentParagraphPropertySourceContract,
  getParagraphPropertySource,
  getParagraphPropertySourceToken,
  getProseDocumentParagraphPropertySourceContract,
  getProseParagraphPropertySourceToken,
  paragraphPropertySourceBelongsToDocument,
  paragraphPropertySourceTokenMatchesContract,
} from "../../packages/core/src/docx/paragraphPropertySource";
import {
  type BrowserDragTarget,
  crossParagraphDragRange,
} from "../visual/browserDragTarget";
import { waitForCanonicalLoadOwner } from "./canonicalLoadOwner";
import { canonicalTextSelection } from "./canonicalTextSelection";

/**
 * Each conversion salts the ids it mints for text box groups and anchors, so
 * two conversions of one document differ there. Compare with each salt named
 * by its order of appearance, which keeps every anchor tied to its box.
 */
const MINTED_ID_ATTRS = new Set(["_docxGroupId", "_docxAnchorId", "anchorId"]);
export const withoutMintedIds = (value: unknown): unknown => {
  const salts = new Map<string, string>();
  return JSON.parse(
    JSON.stringify(value, (key, field: unknown) => {
      if (!MINTED_ID_ATTRS.has(key) || typeof field !== "string") return field;
      const separator = field.indexOf(":");
      if (separator < 0) return field;
      const salt = field.slice(0, separator);
      const name = salts.get(salt) ?? `salt${salts.size}`;
      salts.set(salt, name);
      return `${name}${field.slice(separator)}`;
    }),
  );
};

type CanonicalPlaygroundRef = {
  getDocument: () => Document | null;
  getEditor: () => FolioEditor | null;
  ensureEditorView: () => void;
  loadDocumentBuffer: (buffer: Uint8Array) => Promise<void>;
  save: () => Promise<ArrayBuffer | null>;
  createAIEditSnapshot: () => FolioAIEditSnapshot | null;
  applyDocumentOperations: (options: {
    undefinedReferences: "refuse";
    snapshot: FolioAIEditSnapshot;
    batch: FolioDocumentOperationBatch;
  }) => FolioDocumentOperationResult;
  getSuggestions: () => readonly { suggestionId: string }[];
  exportPendingSuggestions: () => readonly unknown[];
  acceptSuggestion: (
    suggestionId: string,
    options?: { author?: string },
  ) => { accepted: boolean };
  rejectSuggestion: (suggestionId: string) => boolean;
};

export type CanonicalHyperlinkAction =
  | {
      type: "setHyperlink";
      from: number;
      to: number;
      href: string;
      tooltip?: string;
    }
  | { type: "removeHyperlink"; from: number; to: number }
  | {
      type: "insertHyperlink";
      from: number;
      to: number;
      text: string;
      href: string;
      tooltip?: string;
    };

/** The undo handle of the last story edit `replaceStoryBlock` made. */
let lastStoryUndoHandle: FolioDocumentOperationUndoHandle | null = null;

/** The public story handle an operation story names. */
const publicStoryOf = (story: Exclude<OpStory, "main">) =>
  "rId" in story
    ? { type: story.kind, relationshipId: story.rId }
    : { type: story.kind, noteId: story.id };

/** Private interaction-test bridge shared by both playgrounds. */
export const buildCanonicalBridge = (
  getRef: () => CanonicalPlaygroundRef | null,
) => ({
  nativeComposing: () => getRef()?.getEditor()?.getView()?.composing ?? null,
  ensureView: () => {
    const ref = getRef();
    if (!ref) return false;
    ref.ensureEditorView();
    return true;
  },
  load: async (bytes: number[]) => {
    const ref = getRef();
    if (!ref) return false;
    const previousOwner = ref.getEditor()?.getCanonicalDocument();
    await ref.loadDocumentBuffer(new Uint8Array(bytes));
    ref.ensureEditorView();
    await waitForCanonicalLoadOwner({
      previousOwner,
      getOwner: () => getRef()?.getEditor()?.getCanonicalDocument(),
      waitFrame: () =>
        new Promise<void>((resolve) => requestAnimationFrame(() => resolve())),
    });
    return true;
  },
  save: async () => {
    const buffer = await getRef()?.save();
    return buffer ? [...new Uint8Array(buffer)] : null;
  },
  canSnapshot: () =>
    getRef()?.getEditor()?.getCanonicalStoryProjection("main") != null,
  committedVersion: () =>
    getRef()?.getEditor()?.captureCanonicalSave()?.version ?? null,
  snapshot: () => {
    const ref = getRef();
    const editor = ref?.getEditor();
    const state = editor?.getState();
    const canonical = editor?.getCanonicalDocument();
    // Projection and private source ownership must be checked before JSON transport drops the bindings.
    const canonicalProjection = canonical
      ? toProseDoc(
          canonical,
          canonical.package.styles
            ? { styles: canonical.package.styles }
            : undefined,
        )
      : null;
    const sourceContract = canonical
      ? getDocumentParagraphPropertySourceContract(canonical)
      : undefined;
    const content = canonical?.package.document.content ?? [];
    const capturedParagraphCount = content.filter(
      (block) =>
        block.type === "paragraph" &&
        getParagraphPropertySource(block) !== undefined,
    ).length;
    const provenanceValid = Boolean(
      canonical &&
      state &&
      sourceContract &&
      sourceContract ===
        getProseDocumentParagraphPropertySourceContract(state.doc) &&
      content.length > 0 &&
      content.length === state.doc.childCount &&
      content.every((block, index) => {
        if (block.type !== "paragraph") return false;
        const token = getParagraphPropertySourceToken(block);
        const node = state.doc.child(index);
        return token === undefined
          ? getProseParagraphPropertySourceToken(node) == null &&
              getParagraphPropertySource(block) === undefined
          : paragraphPropertySourceTokenMatchesContract(
              token,
              sourceContract,
            ) &&
              getProseParagraphPropertySourceToken(node) === token &&
              (getParagraphPropertySource(block) === undefined ||
                paragraphPropertySourceBelongsToDocument(block, canonical));
      }),
    );
    return {
      active: canonical !== null && canonical !== undefined,
      composing: editor?.getView()?.composing ?? null,
      document: ref?.getDocument() ?? null,
      projectionJSON: state?.doc.toJSON() ?? null,
      canonicalProjectionJSON: canonicalProjection?.toJSON() ?? null,
      projectionMatchesCanonical: Boolean(
        state &&
        canonicalProjection &&
        JSON.stringify(withoutMintedIds(state.doc.toJSON())) ===
          JSON.stringify(withoutMintedIds(canonicalProjection.toJSON())),
      ),
      provenance: { valid: provenanceValid, capturedParagraphCount },
      text: state?.doc.textContent ?? null,
      textSelection: state ? canonicalTextSelection(state) : null,
      selection: state
        ? { from: state.selection.from, to: state.selection.to }
        : null,
      selectionJSON: state?.selection.toJSON() ?? null,
      canUndo: editor?.canUndo() ?? false,
      canRedo: editor?.canRedo() ?? false,
    };
  },
  selectStructuralTarget: (target: BrowserDragTarget) => {
    const view = getRef()?.getEditor()?.getView();
    if (!view) return null;
    if (target === "paragraphs") {
      const range = crossParagraphDragRange(view.state.doc);
      if (!range) return null;
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(view.state.doc, range.from, range.to),
        ),
      );
      view.focus();
      return view.state.selection.toJSON();
    }
    const targets: {
      pos: number;
      size: number;
      type: "paragraph" | "inline";
    }[] = [];
    const name = {
      table: "tableCell",
      list: "paragraph",
      note: "footnoteRef",
      field: "field",
      inlineObject: "image",
    }[target];
    view.state.doc.descendants((node, pos) => {
      if (targets.length >= 2) return false;
      const matches =
        node.type.name === name &&
        (target !== "list" || node.attrs["numPr"] != null);
      if (
        matches ||
        (target === "note" &&
          node.marks.some((mark) => mark.type.name === name))
      )
        targets.push({
          pos,
          size: node.nodeSize,
          type: node.type.name === "paragraph" ? "paragraph" : "inline",
        });
      return target !== "table" || !matches;
    });
    const first = targets.at(0);
    const last = targets.at(1) ?? first;
    if (!first || !last) return null;
    if (target === "table") {
      if (!applyCellSelection(view, first.pos, last.pos)) return null;
    } else {
      const from = Math.max(
        view.state.doc.resolve(first.pos).start(),
        first.type === "paragraph" ? first.pos + 1 : first.pos - 1,
      );
      const to = Math.min(
        view.state.doc.content.size - 1,
        last.type === "paragraph"
          ? last.pos + last.size - 1
          : view.state.doc.resolve(last.pos).end(),
        last.type === "paragraph"
          ? last.pos + Math.min(2, last.size - 1)
          : last.pos + last.size + 2,
      );
      const selection = TextSelection.create(view.state.doc, from, to);
      view.dispatch(view.state.tr.setSelection(selection));
    }
    view.focus();
    return view.state.selection.toJSON();
  },
  setMode: (mode: "editing" | "suggesting") =>
    getRef()
      ?.getEditor()
      ?.setCanonicalMode(
        mode === "suggesting"
          ? { type: "suggesting", author: "Canonical test author" }
          : { type: "editing" },
      ) ?? false,
  resolveRevisions: (
    revisionIds: readonly number[],
    resolution: "accept" | "reject",
  ) =>
    getRef()?.getEditor()?.resolveCanonicalRevisions(revisionIds, resolution) ??
    false,
  /**
   * Replace a secondary story's first paragraph through the public operation
   * API, as a host's batch would, whatever editor has focus.
   */
  replaceStoryBlock: (story: Exclude<OpStory, "main">, text: string) => {
    const editor = getRef()?.getEditor();
    const projection = editor?.getCanonicalStoryProjection(story);
    if (!editor || !projection) return null;
    const snapshot = createFolioAIEditSnapshot(projection);
    const block = snapshot.blocks.at(0);
    if (!block) return null;
    const result = editor.applyCanonicalDocumentOperations({
      undefinedReferences: "refuse",
      snapshot,
      story: publicStoryOf(story),
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [
          { id: "story-edit", type: "replaceBlock", blockId: block.id, text },
        ],
      },
    });
    if (!result) return null;
    lastStoryUndoHandle = result.undoHandle;
    return {
      applied: result.applied.length,
      skipped: result.skipped.map(({ reason }) => reason),
    };
  },
  /** Undo the last story edit `replaceStoryBlock` made. */
  undoStoryBlock: () => {
    const handle = lastStoryUndoHandle;
    if (!handle) return null;
    return (
      getRef()?.getEditor()?.undoCanonicalDocumentOperations(handle)?.status ??
      null
    );
  },
  /** Suggest replacing `find` with `replace` in a block, as a host's batch would. */
  suggestReplace: (
    id: string,
    blockId: string,
    find: string,
    replace: string,
  ) => {
    const ref = getRef();
    const snapshot = ref?.createAIEditSnapshot();
    if (!ref || !snapshot) return null;
    const result = ref.applyDocumentOperations({
      undefinedReferences: "refuse",
      snapshot,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "suggested",
        operations: [{ id, type: "replaceInBlock", blockId, find, replace }],
      },
    });
    return {
      applied: result.applied.map(({ suggestionId }) => suggestionId ?? null),
      skipped: result.skipped.map(({ reason }) => reason),
    };
  },
  /** Propose a comment on a block as a suggestion of its own. */
  suggestBlockComment: (id: string, blockId: string, text: string) => {
    const ref = getRef();
    const snapshot = ref?.createAIEditSnapshot();
    if (!ref || !snapshot) return null;
    const result = ref.applyDocumentOperations({
      undefinedReferences: "refuse",
      snapshot,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "suggested",
        operations: [
          { id, type: "commentOnBlock", blockId, comment: { text } },
        ],
      },
    });
    return {
      applied: result.applied.map(({ suggestionId }) => suggestionId ?? null),
      skipped: result.skipped.map(({ reason }) => reason),
    };
  },
  /** Load a package in whichever session the page runs. */
  loadAnySession: async (bytes: number[]) => {
    const ref = getRef();
    if (!ref) return false;
    await ref.loadDocumentBuffer(new Uint8Array(bytes));
    ref.ensureEditorView();
    return true;
  },
  suggestionIds: () =>
    getRef()
      ?.getSuggestions()
      .map(({ suggestionId }) => suggestionId) ?? null,
  pendingSuggestionCount: () =>
    getRef()?.exportPendingSuggestions().length ?? null,
  resolveSuggestion: (
    suggestionId: string,
    resolution: "accept" | "reject",
  ) => {
    const ref = getRef();
    if (!ref) return false;
    return resolution === "accept"
      ? ref.acceptSuggestion(suggestionId, { author: "Adopter" }).accepted
      : ref.rejectSuggestion(suggestionId);
  },
  select: (from: number, to = from) => {
    const editor = getRef()?.getEditor();
    if (!editor?.getView()) return false;
    editor.setSelection(from, to);
    editor.focus();
    return true;
  },
  /** Run a registered editor command by name, as a toolbar control does. */
  executeNamedCommand: (name: string, args: readonly unknown[]) => {
    const editor = getRef()?.getEditor();
    if (!editor) return false;
    const execute = editor.executeCommand.bind(editor);
    const create = singletonManager.requireCommand(name as never) as (
      ...commandArgs: readonly unknown[]
    ) => Parameters<typeof execute>[0];
    return execute(create(...args));
  },
  executeHyperlink: (action: CanonicalHyperlinkAction) => {
    const editor = getRef()?.getEditor();
    if (!editor) return false;
    editor.setSelection(action.from, action.to);
    switch (action.type) {
      case "setHyperlink":
        return editor.executeCommand(
          singletonManager.requireCommand("setHyperlink")(
            action.href,
            action.tooltip,
          ),
        );
      case "removeHyperlink":
        return editor.executeCommand(
          singletonManager.requireCommand("removeHyperlink")(),
        );
      case "insertHyperlink":
        return editor.executeCommand(
          singletonManager.requireCommand("insertHyperlink")(
            action.text,
            action.href,
            action.tooltip,
          ),
        );
      default: {
        const unreachable: never = action;
        return unreachable;
      }
    }
  },
});

declare global {
  var __folioCanonicalReady: boolean | undefined;
  var __folioCanonical: ReturnType<typeof buildCanonicalBridge> | undefined;
}
