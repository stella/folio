/**
 * A headless editor for tests that drive the editor the way a user does —
 * commands, keys, typed text, paste — and then check what the document model
 * says about the result.
 *
 * The state is assembled the way the editor assembles it: the full extension
 * set's plugins (history, keymaps, input rules, paraId allocation…), the
 * document styles and numbering plugins, and the suggestion-mode plugin in
 * front of them. {@link HeadlessEditorView} implements the part of
 * `EditorView` the plugins' props reach for (`state`, `dispatch`, `someProp`,
 * `endOfTextblock`), so a key press runs the same `handleKeyDown` chain and
 * typed text the same `handleTextInput` chain as in a browser.
 */

import type { Node as PMNode, Slice } from "prosemirror-model";
import { AllSelection, EditorState as PMEditorState, TextSelection } from "prosemirror-state";
import type { EditorState, Plugin, Transaction } from "prosemirror-state";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshotWithStyleResolver } from "../ai-edits/snapshot";
import { assertValidFolioDocumentModel } from "../docx/modelValidation";
import { parseDocx } from "../docx/parser";
import { repackDocx } from "../docx/rezip";
import { docxToMarkdown } from "../docx/server/docxToMarkdown";
import { toMarkdown } from "../markdown";
import { acceptAllChanges, rejectAllChanges } from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { ExtensionManager } from "../prosemirror/extensions/ExtensionManager";
import { ensureBaseDirectionInState } from "../prosemirror/extensions/features/AutoBidiDetectionExtension";
import { ensureParaIdsInDoc } from "../prosemirror/extensions/features/ParaIdAllocatorExtension";
import { createDocumentNumberingPlugin } from "../prosemirror/plugins/documentNumbering";
import {
  createDocumentStylesPlugin,
  getDocumentStyleResolver,
} from "../prosemirror/plugins/documentStyles";
import { createSuggestionModePlugin } from "../prosemirror/plugins/suggestionMode";
import { dispatchEditorTextInput } from "../prosemirror/textInput";
import { singletonManager } from "../prosemirror/schema";
import { createStarterKit } from "../prosemirror/extensions/StarterKit";
import type { Document } from "../types/document";

export type EditorMode = "editing" | "suggesting";

export const EDITOR_MODES: readonly EditorMode[] = ["editing", "suggesting"];

export const HARNESS_AUTHOR = "Conformance";

// ============================================================================
// HEADLESS VIEW
// ============================================================================

type PropHandler = (...args: never[]) => unknown;

const isMacPlatform =
  typeof navigator !== "undefined" && /Mac|iP(hone|[oa]d)/u.test(navigator.platform);

/** The `KeyboardEvent` fields `prosemirror-keymap` reads, for a binding like `Mod-Shift-z`. */
const keyboardEventFor = (binding: string) => {
  const parts = binding.split(/-(?!$)/u);
  const key = parts.at(-1) ?? binding;
  const modifiers = new Set(parts.slice(0, -1));
  const mod = modifiers.has("Mod");
  return {
    key,
    code: key,
    keyCode: 0,
    shiftKey: modifiers.has("Shift"),
    altKey: modifiers.has("Alt"),
    ctrlKey: modifiers.has("Ctrl") || (mod && !isMacPlatform),
    metaKey: modifiers.has("Meta") || (mod && isMacPlatform),
    isComposing: false,
    defaultPrevented: false,
    preventDefault() {
      this.defaultPrevented = true;
    },
  };
};

export class HeadlessEditorView {
  state: EditorState;
  readonly transactions: Transaction[] = [];
  readonly composing = false;
  readonly editable = true;

  constructor(state: EditorState) {
    this.state = state;
  }

  dispatch = (tr: Transaction): void => {
    this.transactions.push(tr);
    this.state = this.state.apply(tr);
  };

  /** `EditorView.someProp` over the state's plugins, in plugin order. */
  someProp(propName: string, f?: (handler: PropHandler) => unknown): unknown {
    for (const plugin of this.state.plugins as readonly Plugin[]) {
      const prop = (plugin.props as Record<string, unknown>)[propName];
      if (prop === undefined || prop === null) {
        continue;
      }
      const value = f ? f(prop as PropHandler) : prop;
      if (value) {
        return value;
      }
    }
    return undefined;
  }

  /** Model-only approximation of the view's line-aware check. */
  endOfTextblock(direction: string, state: EditorState = this.state): boolean {
    const { $head } = state.selection;
    if (!$head.parent.isTextblock) {
      return false;
    }
    const backward = direction === "backward" || direction === "left" || direction === "up";
    return backward ? $head.parentOffset === 0 : $head.parentOffset === $head.parent.content.size;
  }

  /** Run the `handleKeyDown` chain; returns whether a handler claimed the key. */
  pressKey(binding: string): boolean {
    const event = keyboardEventFor(binding);
    const handled = Boolean(
      this.someProp("handleKeyDown", (handler) =>
        (handler as unknown as (view: unknown, event: unknown) => boolean)(this, event),
      ),
    );
    if (handled) {
      return true;
    }
    // What the browser does with a key no handler claimed, for the keys that
    // edit text natively.
    const { selection } = this.state;
    if (event.key === "Backspace" || event.key === "Delete") {
      if (!selection.empty) {
        this.dispatch(this.state.tr.deleteSelection());
        return true;
      }
      const { $head } = selection;
      const backward = event.key === "Backspace";
      if (backward ? $head.parentOffset > 0 : $head.parentOffset < $head.parent.content.size) {
        const from = backward ? $head.pos - 1 : $head.pos;
        this.dispatch(this.state.tr.delete(from, from + 1));
        return true;
      }
    }
    return false;
  }

  /** Type text one character at a time through the text-input funnel. */
  typeText(text: string): void {
    for (const character of text) {
      dispatchEditorTextInput(this as never, character);
    }
  }

  /** Paste a slice through `transformPasted` and `handlePaste`, like a clipboard paste. */
  paste(slice: Slice): void {
    let transformed = slice;
    for (const plugin of this.state.plugins as readonly Plugin[]) {
      const transform = plugin.props.transformPasted;
      if (transform) {
        transformed = transform.call(plugin, transformed, this as never, false);
      }
    }
    const event = { clipboardData: null, preventDefault() {} };
    const handled = this.someProp("handlePaste", (handler) =>
      (handler as unknown as (view: unknown, event: unknown, slice: Slice) => boolean)(
        this,
        event,
        transformed,
      ),
    );
    if (!handled) {
      this.dispatch(
        this.state.tr
          .replaceSelection(transformed)
          .scrollIntoView()
          .setMeta("paste", true)
          .setMeta("uiEvent", "paste"),
      );
    }
  }
}

// ============================================================================
// SESSIONS
// ============================================================================

let sharedManager: ExtensionManager | null = null;

/** One extension manager for every harness state, built like the editor's. */
export const harnessManager = (): ExtensionManager => {
  if (!sharedManager) {
    // The singleton schema must exist before a second manager is built from
    // the starter kit (the kit's extensions read it back at runtime).
    void singletonManager;
    // The harness reports an invalid document itself, after the operation, so
    // it builds its runtime without the per-transaction invariant plugin the
    // test preload installs everywhere else.
    const invariants = globalThis.__folioTransactionInvariants;
    globalThis.__folioTransactionInvariants = undefined;
    try {
      sharedManager = new ExtensionManager(createStarterKit());
      sharedManager.buildSchema();
      sharedManager.initializeRuntime();
    } finally {
      globalThis.__folioTransactionInvariants = invariants;
    }
  }
  return sharedManager;
};

export const parseShapeDocument = (bytes: Uint8Array): Promise<Document> =>
  parseDocx(bytes.slice().buffer, { preloadFonts: false, detectVariables: false });

/** The editor state a mounted editor would hold for `document`. */
export const createHarnessState = (
  document: Document,
  mode: EditorMode,
  extraPlugins: readonly Plugin[] = [],
): EditorState =>
  // The editor's hidden-state assembly (controller/hiddenEditorManager), which
  // tests outside the controller may not import: host plugins first, then the
  // extension runtime, then the document's styles and numbering.
  ensureBaseDirectionInState(
    PMEditorState.create({
      doc: ensureParaIdsInDoc(toProseDoc(document)),
      plugins: [
        ...extraPlugins,
        createSuggestionModePlugin(mode === "suggesting", HARNESS_AUTHOR),
        ...harnessManager().getPlugins(),
        createDocumentStylesPlugin(document.package.styles),
        createDocumentNumberingPlugin(document.package.numbering),
      ],
    }),
  );

// ============================================================================
// POSITIONS
// ============================================================================

export type TextblockMatch = { node: PMNode; pos: number };

/** The first textblock whose text contains `text`. */
export const findTextblock = (doc: PMNode, text: string): TextblockMatch | null => {
  let found: TextblockMatch | null = null;
  doc.descendants((node, pos) => {
    if (found) {
      return false;
    }
    if (node.isTextblock && node.textContent.includes(text)) {
      found = { node, pos };
      return false;
    }
    return true;
  });
  return found;
};

/** Every textblock in document order. */
export const textblocks = (doc: PMNode): TextblockMatch[] => {
  const found: TextblockMatch[] = [];
  doc.descendants((node, pos) => {
    if (node.isTextblock) {
      found.push({ node, pos });
      return false;
    }
    return true;
  });
  return found;
};

export type SelectionPlacement =
  | "caret-start"
  | "caret-middle"
  | "caret-end"
  | "word"
  | "paragraph"
  | "cross-paragraph"
  | "document";

export const SELECTION_PLACEMENTS: readonly SelectionPlacement[] = [
  "caret-start",
  "caret-middle",
  "caret-end",
  "word",
  "paragraph",
  "cross-paragraph",
  "document",
];

/**
 * Place the selection relative to the textblock holding `focus`. Returns null
 * when the placement has no meaning here (no following paragraph to cross into).
 */
export const placeSelection = (
  state: EditorState,
  focus: string,
  placement: SelectionPlacement,
): EditorState | null => {
  const target = findTextblock(state.doc, focus);
  if (!target) {
    throw new Error(`No textblock holds "${focus}"`);
  }
  const start = target.pos + 1;
  const end = target.pos + 1 + target.node.content.size;
  const middle = start + Math.floor(target.node.content.size / 2);
  const create = (anchor: number, head: number) =>
    state.apply(
      state.tr.setSelection(
        TextSelection.between(state.doc.resolve(anchor), state.doc.resolve(head)),
      ),
    );
  switch (placement) {
    case "caret-start": {
      return create(start, start);
    }
    case "caret-middle": {
      return create(middle, middle);
    }
    case "caret-end": {
      return create(end, end);
    }
    case "word": {
      const text = target.node.textContent;
      const wordEnd = text.search(/\s/u);
      const length = wordEnd <= 0 ? Math.min(3, text.length) : wordEnd;
      return length === 0 ? null : create(start, start + length);
    }
    case "paragraph": {
      return end > start ? create(start, end) : null;
    }
    case "cross-paragraph": {
      const blocks = textblocks(state.doc);
      const index = blocks.findIndex((block) => block.pos === target.pos);
      const next = blocks[index + 1];
      if (!next) {
        return null;
      }
      const nextMiddle = next.pos + 1 + Math.floor(next.node.content.size / 2);
      return create(middle, nextMiddle);
    }
    case "document": {
      return state.apply(state.tr.setSelection(new AllSelection(state.doc)));
    }
    default: {
      const unhandled: never = placement;
      throw new Error(`Unhandled placement ${String(unhandled)}`);
    }
  }
};

// ============================================================================
// OBSERVATIONS
// ============================================================================

/** What a reader of the document sees: one entry per content block. */
export type ContentSummary = readonly Record<string, unknown>[];

export const summarizeState = (state: EditorState): ContentSummary =>
  createFolioAIEditSnapshotWithStyleResolver(state.doc, getDocumentStyleResolver(state)).blocks.map(
    stripBlockIdentity,
  );

export const summarizeReviewer = (reviewer: FolioDocxReviewer): ContentSummary =>
  reviewer.getContent().map(stripBlockIdentity);

export const stripBlockIdentity = (block: object): Record<string, unknown> => {
  const {
    id: _id,
    idStability: _stability,
    containerPath,
    ...rest
  } = block as Record<string, unknown>;
  return containerPath === undefined
    ? rest
    : {
        ...rest,
        containerPath: (containerPath as readonly Record<string, unknown>[]).map(
          ({ id: _containerId, ...entry }) => entry,
        ),
      };
};

export type SavedDocument = {
  model: Document;
  bytes: Uint8Array;
};

/** Convert the editor state to the model, validate it, and repack it. */
export const saveHarnessState = async (
  state: EditorState,
  base: Document,
): Promise<SavedDocument> => {
  const model = fromProseDoc(state.doc, base);
  assertValidFolioDocumentModel(model, "Editor state converts to an invalid DOCX model");
  const bytes = new Uint8Array(await repackDocx(model, { updateModifiedDate: false }));
  return { model, bytes };
};

const EFFECTIVE_PARAGRAPH_ATTRS = [
  "styleId",
  "alignment",
  "indentLeft",
  "indentRight",
  "spaceBefore",
  "spaceAfter",
  "lineSpacing",
  "lineSpacingRule",
] as const;

/**
 * What the editor paints for each paragraph: its effective (cascaded)
 * properties, as the editor state holds them. The reader snapshot reports
 * direct formatting only, so an edit that paints but does not save, or saves
 * as something else, shows up here.
 */
export const summarizeEffectiveParagraphs = (state: EditorState): ContentSummary =>
  textblocks(state.doc).map(({ node }) => {
    const entry: Record<string, unknown> = { text: node.textContent.slice(0, 40) };
    for (const key of EFFECTIVE_PARAGRAPH_ATTRS) {
      const value: unknown = node.attrs[key];
      if (value !== null && value !== undefined) {
        entry[key] = value;
      }
    }
    const firstLine: unknown = node.attrs["indentFirstLine"];
    if (typeof firstLine === "number") {
      entry["indentFirstLine"] =
        node.attrs["hangingIndent"] === true ? -Math.abs(firstLine) : firstLine;
    }
    const numPr = node.attrs["numPr"] as { numId?: number; ilvl?: number } | null | undefined;
    if (numPr?.numId !== undefined) {
      entry["numbering"] = `${numPr.numId}:${numPr.ilvl ?? 0}`;
    }
    return entry;
  });

export type Readback = {
  summary: ContentSummary;
  effective: ContentSummary;
  markdown: string;
};

/**
 * Read saved bytes back the way a host would: reviewer blocks, Markdown, and
 * what an editor reopening them paints.
 */
export const readBack = async (bytes: Uint8Array): Promise<Readback> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer);
  const reopened = createHarnessState(await parseShapeDocument(bytes), "editing");
  return {
    summary: summarizeReviewer(reviewer),
    effective: summarizeEffectiveParagraphs(reopened),
    markdown: await docxToMarkdown(bytes.slice().buffer),
  };
};

export const modelMarkdown = (model: Document): string => toMarkdown(model);

/** Resolve every tracked change through the editor's accept-all / reject-all command. */
export const resolveAllChanges = (state: EditorState, mode: "accept" | "reject"): EditorState => {
  let resolved = state;
  (mode === "accept" ? acceptAllChanges() : rejectAllChanges())(state, (tr) => {
    resolved = resolved.apply(tr);
  });
  return resolved;
};
