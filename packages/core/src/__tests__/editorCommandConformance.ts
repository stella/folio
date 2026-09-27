/**
 * The editor-command conformance matrix: every command, key binding, typed
 * marker and paste the editor offers, applied to every document shape in both
 * editing and suggesting mode, and then checked against what a host does next
 * — save, reopen, undo, and resolve tracked changes.
 *
 * A case passes when
 * (a) the edited state converts to a document model that validates,
 * (b) `repackDocx` writes it,
 * (c) reading the saved bytes back yields the blocks (kinds, labels, text,
 *     formatting) and the Markdown the edited state showed,
 * (d) undo returns the original document and redo the edited one, and
 * (e) in suggesting mode the operation either applies as it does in editing
 *     mode or is refused in editing mode too; rejecting every change returns
 *     what rejecting them in the original returned, and accepting every change
 *     yields what accepting them after the editing-mode run yields.
 *
 * Refusals are allowed — a table command outside a table does nothing — but a
 * refusal that only happens in suggesting mode is a silent no-op, and is
 * reported. Known gaps are recorded in `editorCommandConformance.known.ts`.
 */

import { Fragment, Slice } from "prosemirror-model";
import type { Node as PMNode } from "prosemirror-model";
import { redo, undo } from "prosemirror-history";
import type { EditorState } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import type { DocumentShape } from "./documentShapes";
import {
  createHarnessState,
  EDITOR_MODES,
  harnessManager,
  HeadlessEditorView,
  modelMarkdown,
  parseShapeDocument,
  placeSelection,
  readBack,
  resolveAllChanges,
  saveHarnessState,
  summarizeEffectiveParagraphs,
  summarizeState,
  textblocks,
} from "./editorHarness";
import type { ContentSummary, EditorMode, SelectionPlacement } from "./editorHarness";
import { getCachedNumberingMap } from "../docx/numberingParser";
import {
  acceptAllChanges,
  acceptChange,
  addCommentMark,
  rejectAllChanges,
  rejectChange,
} from "../prosemirror/commands/comments";
import { clearFormatting } from "../prosemirror/commands/formatting";
import { insertPageBreak } from "../prosemirror/commands/pageBreak";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { deleteSelectionAsSuggestion } from "../prosemirror/plugins/suggestionMode";
import type { ResolvedStyleAttrs } from "../prosemirror/extensions/core/ParagraphExtension";
import { createStyleResolver } from "../prosemirror/styles/styleResolver";
import type { Comment, Document } from "../types/document";

// ============================================================================
// OPERATIONS
// ============================================================================

export type OperationContext = {
  view: HeadlessEditorView;
  /** The model the host saves against. Host operations may add records to it. */
  base: Document;
  /** Text of the shape's focus paragraph. */
  focus: string;
};

export type ConformanceOperation = {
  /** Stable identifier, e.g. `command:toggleBold` or `key:Enter`. */
  id: string;
  /** Placements the default tier tries; the full tier tries every placement. */
  placements: readonly SelectionPlacement[];
  /**
   * How suggesting mode records the operation. `direct` operations are not
   * revisions in OOXML either (a comment, resolving a change), so rejecting
   * every change is not expected to undo them.
   */
  suggesting?: "tracked" | "direct";
  /**
   * Run the operation. Returns the command's verdict when it has one: `false`
   * is a refusal. Anything else is judged by whether the document changed.
   */
  run: (context: OperationContext) => boolean | undefined;
};

const CARET: readonly SelectionPlacement[] = ["caret-middle"];
const RANGE: readonly SelectionPlacement[] = ["word"];

const registryCommand = (
  name: string,
  args: readonly unknown[] = [],
  {
    variant,
    placements = CARET,
  }: { variant?: string; placements?: readonly SelectionPlacement[] } = {},
): ConformanceOperation => ({
  id: variant === undefined ? `command:${name}` : `command:${name}(${variant})`,
  placements,
  run: ({ view }) =>
    harnessManager().requireCommand(name)(...args)(view.state, view.dispatch, view as never),
});

/** Apply a paragraph style the way the toolbar does: with the resolved style attrs. */
const applyStyleLikeHost = (styleId: string): ConformanceOperation => ({
  id: `command:applyStyle(${styleId})`,
  placements: CARET,
  run: ({ view, base }) => {
    const applyStyle = harnessManager().requireCommand("applyStyle");
    const styles = base.package.styles;
    if (!styles) {
      return applyStyle(styleId)(view.state, view.dispatch, view as never);
    }
    const resolver = createStyleResolver(styles);
    const resolved = resolver.resolveParagraphStyle(styleId);
    const attrs: ResolvedStyleAttrs = {
      numbering: base.package.numbering ? getCachedNumberingMap(base.package.numbering) : null,
    };
    if (resolved.paragraphFormatting) {
      attrs.paragraphFormatting = resolved.paragraphFormatting;
    }
    if (resolved.runFormatting) {
      attrs.runFormatting = resolved.runFormatting;
    }
    const styleName = resolver.getStyle(styleId)?.name;
    if (styleName) {
      attrs.styleName = styleName;
    }
    return applyStyle(styleId, attrs)(view.state, view.dispatch, view as never);
  },
});

const BORDER = { style: "single", size: 8, color: { rgb: "FF0000" } } as const;

/**
 * Every registry command, with the arguments the matrix calls it with, or the
 * reason it is not driven directly. A command registered without an entry
 * here fails the coverage check, so a new command joins the matrix or states
 * why it cannot.
 */
export const REGISTRY_COMMAND_OPERATIONS: Readonly<
  Record<string, readonly ConformanceOperation[] | { excluded: string }>
> = {
  // Marks
  toggleBold: [registryCommand("toggleBold", [], { placements: RANGE })],
  toggleItalic: [registryCommand("toggleItalic", [], { placements: RANGE })],
  toggleUnderline: [registryCommand("toggleUnderline", [], { placements: RANGE })],
  toggleStrike: [registryCommand("toggleStrike", [], { placements: RANGE })],
  toggleSuperscript: [registryCommand("toggleSuperscript", [], { placements: RANGE })],
  toggleSubscript: [registryCommand("toggleSubscript", [], { placements: RANGE })],
  setTextColor: [registryCommand("setTextColor", [{ rgb: "C00000" }], { placements: RANGE })],
  clearTextColor: [registryCommand("clearTextColor", [], { placements: RANGE })],
  setHighlight: [registryCommand("setHighlight", ["yellow"], { placements: RANGE })],
  clearHighlight: [registryCommand("clearHighlight", [], { placements: RANGE })],
  setFontSize: [registryCommand("setFontSize", [28], { placements: RANGE })],
  clearFontSize: [registryCommand("clearFontSize", [], { placements: RANGE })],
  setFontFamily: [registryCommand("setFontFamily", ["Arial"], { placements: RANGE })],
  clearFontFamily: [registryCommand("clearFontFamily", [], { placements: RANGE })],
  setUnderlineStyle: [registryCommand("setUnderlineStyle", ["double"], { placements: RANGE })],
  setHyperlink: [registryCommand("setHyperlink", ["https://example.org/"], { placements: RANGE })],
  removeHyperlink: [registryCommand("removeHyperlink", [], { placements: RANGE })],
  insertHyperlink: [registryCommand("insertHyperlink", ["a link", "https://example.org/"])],
  insertFootnote: [registryCommand("insertFootnote", [7])],
  insertEndnote: [registryCommand("insertEndnote", [7])],
  deleteNoteRef: [registryCommand("deleteNoteRef", [], { placements: ["paragraph"] })],

  // Paragraph
  setAlignment: [registryCommand("setAlignment", ["center"], { variant: "center" })],
  alignLeft: [registryCommand("alignLeft")],
  alignCenter: [registryCommand("alignCenter")],
  alignRight: [registryCommand("alignRight")],
  alignJustify: [registryCommand("alignJustify")],
  setLineSpacing: [registryCommand("setLineSpacing", [360, "auto"], { variant: "360" })],
  singleSpacing: [registryCommand("singleSpacing")],
  oneAndHalfSpacing: [registryCommand("oneAndHalfSpacing")],
  doubleSpacing: [registryCommand("doubleSpacing")],
  setSpaceBefore: [registryCommand("setSpaceBefore", [240], { variant: "240" })],
  setSpaceAfter: [registryCommand("setSpaceAfter", [240], { variant: "240" })],
  increaseIndent: [registryCommand("increaseIndent")],
  decreaseIndent: [registryCommand("decreaseIndent")],
  setIndentLeft: [registryCommand("setIndentLeft", [720], { variant: "720" })],
  setIndentRight: [registryCommand("setIndentRight", [720], { variant: "720" })],
  setIndentFirstLine: [
    registryCommand("setIndentFirstLine", [360], { variant: "360" }),
    registryCommand("setIndentFirstLine", [360, true], { variant: "hanging" }),
  ],
  applyStyle: [applyStyleLikeHost("Heading1"), applyStyleLikeHost("Heading2")],
  clearStyle: [registryCommand("clearStyle")],
  insertSectionBreak: [
    registryCommand("insertSectionBreak", ["nextPage"], { variant: "nextPage" }),
    registryCommand("insertSectionBreak", ["continuous"], { variant: "continuous" }),
  ],
  removeSectionBreak: [registryCommand("removeSectionBreak")],
  generateTOC: [registryCommand("generateTOC", [{ title: "Contents" }])],
  toggleBidi: [registryCommand("toggleBidi")],
  setRtl: [registryCommand("setRtl")],
  setLtr: [registryCommand("setLtr")],
  setTabs: [registryCommand("setTabs", [[{ position: 2880, alignment: "right" }]])],
  addTabStop: [registryCommand("addTabStop", [1440, "left", "dot"])],
  removeTabStop: [registryCommand("removeTabStop", [1440])],

  // Lists
  toggleBulletList: [registryCommand("toggleBulletList")],
  toggleNumberedList: [registryCommand("toggleNumberedList")],
  increaseListLevel: [registryCommand("increaseListLevel")],
  decreaseListLevel: [registryCommand("decreaseListLevel")],
  removeList: [registryCommand("removeList")],
  restartNumbering: [registryCommand("restartNumbering")],
  continueNumbering: [registryCommand("continueNumbering")],
  setNumberingValue: [registryCommand("setNumberingValue", [3], { variant: "3" })],

  // Tables
  insertTable: [registryCommand("insertTable", [2, 3])],
  addRowAbove: [registryCommand("addRowAbove")],
  addRowBelow: [registryCommand("addRowBelow")],
  deleteRow: [registryCommand("deleteRow")],
  addColumnLeft: [registryCommand("addColumnLeft")],
  addColumnRight: [registryCommand("addColumnRight")],
  deleteColumn: [registryCommand("deleteColumn")],
  deleteTable: [registryCommand("deleteTable")],
  selectTable: [registryCommand("selectTable")],
  selectRow: [registryCommand("selectRow")],
  selectColumn: [registryCommand("selectColumn")],
  mergeCells: [registryCommand("mergeCells", [], { placements: ["cross-paragraph"] })],
  splitCell: [registryCommand("splitCell")],
  setCellBorder: [registryCommand("setCellBorder", ["all", BORDER])],
  setTableBorderPreset: [registryCommand("setTableBorderPreset", ["none"])],
  setTableBorders: [registryCommand("setTableBorders", ["outside", BORDER])],
  removeTableBorders: [registryCommand("removeTableBorders")],
  setAllTableBorders: [registryCommand("setAllTableBorders", [BORDER])],
  setOutsideTableBorders: [registryCommand("setOutsideTableBorders", [BORDER])],
  setInsideTableBorders: [registryCommand("setInsideTableBorders", [BORDER])],
  setCellVerticalAlign: [registryCommand("setCellVerticalAlign", ["center"])],
  setCellMargins: [registryCommand("setCellMargins", [{ top: 120, left: 240 }])],
  setCellTextDirection: [registryCommand("setCellTextDirection", ["btLr"])],
  toggleNoWrap: [registryCommand("toggleNoWrap")],
  setRowHeight: [registryCommand("setRowHeight", [600, "atLeast"])],
  toggleHeaderRow: [registryCommand("toggleHeaderRow")],
  distributeColumns: [registryCommand("distributeColumns")],
  autoFitContents: [registryCommand("autoFitContents")],
  setTableProperties: [
    registryCommand("setTableProperties", [
      { width: 5000, widthType: "pct", justification: "center" },
    ]),
  ],
  applyTableStyle: [registryCommand("applyTableStyle", [{ styleId: "TableGrid" }])],
  setCellFillColor: [registryCommand("setCellFillColor", ["FFFF00"])],
  setTableBorderColor: [registryCommand("setTableBorderColor", ["00B050"])],
  setTableBorderWidth: [registryCommand("setTableBorderWidth", [12])],

  // Not driven directly
  undo: { excluded: "Exercised by every case's undo check (d)." },
  redo: { excluded: "Exercised by every case's redo check (d)." },
  extractSelectionContext: {
    excluded: "Reads the selection for a host callback; it never edits the document.",
  },
};

/** Two plain paragraphs, open at both ends, as a clipboard delivers them. */
const PASTED_PARAGRAPHS = (schemaDoc: PMNode) => {
  const { schema } = schemaDoc.type;
  return new Slice(
    Fragment.from([
      schema.node("paragraph", null, schema.text("Pasted first")),
      schema.node("paragraph", null, schema.text("Pasted second")),
    ]),
    1,
    1,
  );
};

const PASTED_TABLE = (schemaDoc: PMNode) => {
  const { schema } = schemaDoc.type;
  const cell = (text: string) =>
    schema.node("tableCell", null, [schema.node("paragraph", null, schema.text(text))]);
  const row = (left: string, right: string) =>
    schema.node("tableRow", null, [cell(left), cell(right)]);
  return new Slice(
    Fragment.from(
      schema.node("table", null, [
        row("First cell", "Second cell"),
        row("Third cell", "Fourth cell"),
      ]),
    ),
    0,
    0,
  );
};

const PASTED_LIST = (schemaDoc: PMNode) => {
  const { schema } = schemaDoc.type;
  const item = (text: string, level: number) =>
    schema.node(
      "paragraph",
      { _pastedHtmlList: { group: 1, kind: "bullet", level } },
      schema.text(text),
    );
  return new Slice(
    Fragment.from([item("First bullet", 0), item("Second bullet", 0), item("Nested bullet", 1)]),
    1,
    1,
  );
};

const keyOperation = (
  binding: string,
  placements: readonly SelectionPlacement[] = CARET,
): ConformanceOperation => ({
  id: `key:${binding}`,
  placements,
  run: ({ view }) => view.pressKey(binding),
});

/**
 * Every key binding the extension keymaps register, driven through the full
 * `handleKeyDown` chain, or the reason it is not. A binding registered without
 * an entry here fails the coverage check.
 */
export const KEY_BINDING_OPERATIONS: Readonly<
  Record<string, readonly ConformanceOperation[] | { excluded: string; macOSOnly?: true }>
> = {
  Enter: [keyOperation("Enter", ["caret-middle", "caret-end", "paragraph"])],
  "Shift-Enter": [keyOperation("Shift-Enter")],
  "Mod-Enter": [keyOperation("Mod-Enter")],
  Backspace: [keyOperation("Backspace", ["caret-start", "caret-middle", "word"])],
  Delete: [keyOperation("Delete", ["caret-end", "cross-paragraph"])],
  "Mod-Backspace": [keyOperation("Mod-Backspace", ["caret-start"])],
  "Mod-Delete": [keyOperation("Mod-Delete", ["caret-end"])],
  "Shift-Backspace": [keyOperation("Shift-Backspace", ["caret-start"])],
  Tab: [keyOperation("Tab", ["caret-start", "caret-middle"])],
  "Shift-Tab": [keyOperation("Shift-Tab", ["caret-start"])],
  "Mod-b": [keyOperation("Mod-b", RANGE)],
  "Mod-i": [keyOperation("Mod-i", RANGE)],
  "Mod-u": [keyOperation("Mod-u", RANGE)],
  "Mod-z": { excluded: "Undo; exercised by every case's undo check (d)." },
  "Mod-y": { excluded: "Redo; exercised by every case's redo check (d)." },
  "Mod-Shift-z": { excluded: "Redo; exercised by every case's redo check (d)." },
  "Mod-Alt-v": {
    excluded:
      "Paste without formatting reads the system clipboard; paste is driven by the paste operations.",
  },
  "Mod-a": { excluded: "Select all; moves the selection only (the `document` placement)." },
  Escape: { excluded: "Selects the parent node; moves the selection only." },
  "Ctrl-a": { excluded: "macOS line-start motion; moves the selection only.", macOSOnly: true },
  "Ctrl-e": { excluded: "macOS line-end motion; moves the selection only.", macOSOnly: true },
  "Ctrl-h": { excluded: "macOS alias of Backspace, which is driven.", macOSOnly: true },
  "Ctrl-d": { excluded: "macOS alias of Delete, which is driven.", macOSOnly: true },
  "Alt-Backspace": { excluded: "macOS alias of Mod-Backspace, which is driven.", macOSOnly: true },
  "Ctrl-Alt-Backspace": {
    excluded: "macOS alias of Mod-Delete, which is driven.",
    macOSOnly: true,
  },
  "Alt-Delete": { excluded: "macOS alias of Mod-Delete, which is driven.", macOSOnly: true },
  "Alt-d": { excluded: "macOS alias of Mod-Delete, which is driven.", macOSOnly: true },
};

const typing = (text: string, label: string): ConformanceOperation => ({
  id: `type:${label}`,
  placements: ["caret-start"],
  run: ({ view }) => {
    view.typeText(text);
    return undefined;
  },
});

const nextCommentId = (base: Document, doc: PMNode): number => {
  let highest = -1;
  for (const comment of base.package.document.comments ?? []) {
    highest = Math.max(highest, comment.id);
  }
  doc.descendants((node) => {
    for (const mark of node.marks) {
      const id: unknown = mark.attrs["commentId"];
      if (typeof id === "number") {
        highest = Math.max(highest, id);
      }
    }
    return true;
  });
  return highest + 1;
};

/**
 * Operations that are not registry commands: typed text and the list
 * autoformat markers, paste, and the host-level flows (comments, accept and
 * reject, page break, clear formatting).
 */
export const EXTRA_OPERATIONS: readonly ConformanceOperation[] = [
  typing("xyz", "text"),
  { ...typing("xyz", "text"), id: "type:text(mid)", placements: CARET },
  typing("- ", "bullet-marker"),
  typing("* ", "star-marker"),
  typing("1. ", "number-marker"),
  {
    id: "paste:plain",
    placements: ["caret-middle", "word"],
    run: ({ view }) => {
      view.paste(new Slice(Fragment.from(view.state.schema.text("Inline paste")), 0, 0));
      return undefined;
    },
  },
  {
    id: "paste:paragraphs",
    placements: ["caret-middle", "word"],
    run: ({ view }) => {
      view.paste(PASTED_PARAGRAPHS(view.state.doc));
      return undefined;
    },
  },
  {
    id: "paste:table",
    placements: ["caret-middle"],
    run: ({ view }) => {
      view.paste(PASTED_TABLE(view.state.doc));
      return undefined;
    },
  },
  {
    id: "paste:list",
    placements: ["caret-middle"],
    run: ({ view }) => {
      view.paste(PASTED_LIST(view.state.doc));
      return undefined;
    },
  },
  {
    // Copy the focus paragraph and the one after it, paste them back: an
    // internal copy carries every paragraph attribute, numbering included.
    id: "paste:copied-blocks",
    placements: ["caret-end"],
    run: ({ view, focus }) => {
      const blocks = textblocks(view.state.doc);
      const index = blocks.findIndex(({ node }) => node.textContent.includes(focus));
      const first = blocks[index];
      const second = blocks[index + 1];
      if (!first || !second) {
        return false;
      }
      view.paste(view.state.doc.slice(first.pos + 1, second.pos + 1 + second.node.content.size));
      return undefined;
    },
  },
  {
    id: "host:cut",
    placements: ["word", "paragraph", "cross-paragraph"],
    run: ({ view }) => {
      if (view.state.selection.empty) {
        return false;
      }
      if (!deleteSelectionAsSuggestion(view.state, view.dispatch)) {
        view.dispatch(view.state.tr.deleteSelection());
      }
      return true;
    },
  },
  {
    id: "host:addComment",
    suggesting: "direct",
    placements: RANGE,
    run: ({ view, base }) => {
      const id = nextCommentId(base, view.state.doc);
      const applied = addCommentMark(id)(view.state, view.dispatch);
      if (applied) {
        const comment: Comment = {
          id,
          author: "Conformance",
          date: "2026-01-01T00:00:00Z",
          content: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "A note." }] }],
            },
          ],
        };
        base.package.document.comments = [...(base.package.document.comments ?? []), comment];
      }
      return applied;
    },
  },
  {
    id: "host:acceptChange",
    suggesting: "direct",
    placements: ["paragraph"],
    run: ({ view }) =>
      acceptChange(view.state.selection.from, view.state.selection.to)(view.state, view.dispatch),
  },
  {
    id: "host:rejectChange",
    suggesting: "direct",
    placements: ["paragraph"],
    run: ({ view }) =>
      rejectChange(view.state.selection.from, view.state.selection.to)(view.state, view.dispatch),
  },
  {
    id: "host:acceptAllChanges",
    suggesting: "direct",
    placements: CARET,
    run: ({ view }) => acceptAllChanges()(view.state, view.dispatch),
  },
  {
    id: "host:rejectAllChanges",
    suggesting: "direct",
    placements: CARET,
    run: ({ view }) => rejectAllChanges()(view.state, view.dispatch),
  },
  {
    id: "host:insertPageBreak",
    placements: CARET,
    run: ({ view }) => insertPageBreak(view.state, view.dispatch),
  },
  {
    id: "host:clearFormatting",
    placements: ["paragraph"],
    run: ({ view }) => clearFormatting(view.state, view.dispatch),
  },
];

export const CONFORMANCE_OPERATIONS: readonly ConformanceOperation[] = [
  ...Object.values(REGISTRY_COMMAND_OPERATIONS).flatMap((entry) =>
    Array.isArray(entry) ? entry : [],
  ),
  ...Object.values(KEY_BINDING_OPERATIONS).flatMap((entry) => (Array.isArray(entry) ? entry : [])),
  ...EXTRA_OPERATIONS,
];

// ============================================================================
// CASES
// ============================================================================

export type ViolationKind =
  | "threw"
  | "invalid-model"
  | "table-grid"
  | "readback-blocks"
  | "readback-painted"
  | "readback-markdown"
  | "undo"
  | "redo"
  | "silent-refusal"
  | "reject-mismatch"
  | "accept-mismatch";

export type Violation = {
  kind: ViolationKind;
  mode: EditorMode;
  detail: string;
};

type LoadedShape = {
  base: Document;
  states: Record<EditorMode, EditorState>;
};

const loadedShapes = new Map<string, Promise<LoadedShape>>();

export const loadShape = (shape: DocumentShape): Promise<LoadedShape> => {
  let pending = loadedShapes.get(shape.id);
  if (!pending) {
    pending = (async () => {
      const base = await parseShapeDocument(await shape.build());
      return {
        base,
        states: {
          editing: createHarnessState(base, "editing"),
          suggesting: createHarnessState(base, "suggesting"),
        },
      };
    })();
    loadedShapes.set(shape.id, pending);
  }
  return pending;
};

/** A base the host may add records to without touching the shared parse. */
const caseBase = (base: Document): Document => ({
  ...base,
  package: {
    ...base.package,
    document: {
      ...base.package.document,
      ...(base.package.document.comments ? { comments: [...base.package.document.comments] } : {}),
    },
  },
});

type ModeRun = {
  mode: EditorMode;
  before: EditorState;
  after: EditorState;
  base: Document;
  status: "refused" | "unchanged" | "changed" | "threw";
  error?: unknown;
};

const runMode = (
  loaded: LoadedShape,
  shape: DocumentShape,
  operation: ConformanceOperation,
  placement: SelectionPlacement,
  mode: EditorMode,
): ModeRun | null => {
  const before = placeSelection(loaded.states[mode], shape.focus, placement);
  if (!before) {
    return null;
  }
  const base = caseBase(loaded.base);
  const view = new HeadlessEditorView(before);
  let verdict: boolean | undefined;
  try {
    verdict = operation.run({ view, base, focus: shape.focus });
  } catch (error) {
    return { mode, before, after: view.state, base, status: "threw", error };
  }
  if (!view.state.doc.eq(before.doc)) {
    return { mode, before, after: view.state, base, status: "changed" };
  }
  return {
    mode,
    before,
    after: view.state,
    base,
    status: verdict === false ? "refused" : "unchanged",
  };
};

const errorText = (error: unknown): string =>
  error instanceof Error ? `${error.name}: ${error.message}` : String(error);

const firstDifference = (left: string, right: string): string => {
  let index = 0;
  while (index < left.length && left[index] === right[index]) {
    index += 1;
  }
  const from = Math.max(0, index - 60);
  return `…${left.slice(from, index + 80)}… ≠ …${right.slice(from, index + 80)}…`;
};

const blockLabel = (block: Record<string, unknown> | undefined): string =>
  JSON.stringify(String(block?.["text"] ?? "").slice(0, 32));

/** The first block that differs, and the fields it differs in. */
const compareSummaries = (expected: ContentSummary, actual: ContentSummary): string | null => {
  if (JSON.stringify(expected) === JSON.stringify(actual)) {
    return null;
  }
  const count = Math.max(expected.length, actual.length);
  for (let index = 0; index < count; index += 1) {
    const left = expected[index];
    const right = actual[index];
    if (JSON.stringify(left) === JSON.stringify(right)) {
      continue;
    }
    if (!left || !right) {
      return `block count ${expected.length} ≠ ${actual.length} (first unmatched ${blockLabel(left ?? right)})`;
    }
    const fields = [...new Set([...Object.keys(left), ...Object.keys(right)])]
      .filter((field) => JSON.stringify(left[field]) !== JSON.stringify(right[field]))
      .map(
        (field) =>
          `${field}: ${JSON.stringify(left[field]) ?? "∅"} ≠ ${JSON.stringify(right[field]) ?? "∅"}`,
      );
    return `block ${index} ${blockLabel(left)}: ${fields.join("; ")}`;
  }
  return null;
};

/**
 * Markdown with the values a save legitimately renumbers taken out: revision
 * ids (normalized package-wide on save), revision dates, and the paragraph id
 * an image's virtual path is named after.
 */
const normalizeMarkdown = (markdown: string): string =>
  markdown
    .replaceAll(/ id="\d+"/gu, ' id="#"')
    .replaceAll(/ date="[^"]*"/gu, ' date="#"')
    .replaceAll(/\.\/images\/[0-9A-F]{8}-img/gu, "./images/#-img")
    .replaceAll(/#_Toc\d+/gu, "#_Toc#")
    // One revision the save wrote as several wrappers (around a hyperlink, say).
    .replaceAll(/<\/(ins|del)><\1 author="[^"]*" date="#" id="#">/gu, "");

const compareText = (expected: string, actual: string): string | null => {
  const left = normalizeMarkdown(expected);
  const right = normalizeMarkdown(actual);
  return left === right ? null : firstDifference(JSON.stringify(left), JSON.stringify(right));
};

/** Where two documents first differ, named by node type and the attributes or marks that differ. */
const describeDocDifference = (expected: PMNode, actual: PMNode): string => {
  const position = expected.content.findDiffStart(actual.content);
  if (position === null) {
    return "document attributes differ";
  }
  const left = expected.resolve(Math.min(position, expected.content.size));
  const right = actual.resolve(Math.min(position, actual.content.size));
  const leftNode = left.nodeAfter ?? left.parent;
  const rightNode = right.nodeAfter ?? right.parent;
  if (leftNode.type !== rightNode.type) {
    return `at ${position}: ${leftNode.type.name} ≠ ${rightNode.type.name}`;
  }
  const attrs = [...new Set([...Object.keys(leftNode.attrs), ...Object.keys(rightNode.attrs)])]
    .filter((key) => JSON.stringify(leftNode.attrs[key]) !== JSON.stringify(rightNode.attrs[key]))
    .map(
      (key) =>
        `${key}: ${JSON.stringify(leftNode.attrs[key])?.slice(0, 120)} ≠ ${JSON.stringify(rightNode.attrs[key])?.slice(0, 120)}`,
    );
  const marks = (node: PMNode) => node.marks.map((mark) => mark.type.name).join(",");
  const markDifference =
    marks(leftNode) === marks(rightNode) ? [] : [`marks: ${marks(leftNode)} ≠ ${marks(rightNode)}`];
  const text =
    leftNode.textContent === rightNode.textContent
      ? []
      : [
          `text: ${JSON.stringify(leftNode.textContent.slice(0, 40))} ≠ ${JSON.stringify(rightNode.textContent.slice(0, 40))}`,
        ];
  return `at ${position} (${leftNode.type.name}): ${[...attrs, ...markDifference, ...text].join("; ") || "content differs"}`;
};

/** One step of `command` (undo or redo), or null when it has nothing to do. */
const historyStep = (state: EditorState, command: typeof undo): EditorState | null => {
  let next: EditorState | null = null;
  command(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
};

/** Apply `command` (undo or redo) until it has nothing left to do. */
const exhaustHistory = (state: EditorState, command: typeof undo): EditorState => {
  let current = state;
  for (let step = 0; step < 50; step += 1) {
    const next = historyStep(current, command);
    if (!next) {
      break;
    }
    current = next;
  }
  return current;
};

type Observation = { summary: ContentSummary; markdown: string };

const observe = (state: EditorState, base: Document): Observation => ({
  summary: summarizeState(state),
  markdown: modelMarkdown(fromProseDoc(state.doc, base)),
});

const observeReopened = async (state: EditorState, base: Document): Promise<Observation> => {
  const { bytes } = await saveHarnessState(state, base);
  const reopened = await readBack(bytes);
  return { summary: reopened.summary, markdown: reopened.markdown };
};

const compareObservations = (expected: Observation, actual: Observation): string | null =>
  compareSummaries(expected.summary, actual.summary) ??
  compareText(expected.markdown, actual.markdown);

/**
 * Every table whose cells no longer tile its grid: a merge left longer than
 * the table, or a row the table fixer had to pad. Such a table saves as
 * `w:vMerge` runs and grid columns its source never had.
 */
const tableGridIssues = (doc: PMNode): string[] => {
  const issues: string[] = [];
  doc.descendants((node, pos) => {
    if (node.type.spec["tableRole"] !== "table") {
      return true;
    }
    const map = TableMap.get(node);
    const columnWidths: unknown = node.attrs["columnWidths"];
    if (map.problems && map.problems.length > 0) {
      issues.push(`table at ${pos}: ${JSON.stringify(map.problems)}`);
    } else if (Array.isArray(columnWidths) && columnWidths.length !== map.width) {
      issues.push(
        `table at ${pos}: ${map.width} cell columns over a ${columnWidths.length}-column grid`,
      );
    }
    return true;
  });
  return issues;
};

/** Checks (a)–(d) for one mode's changed state. */
const checkChangedState = async (run: ModeRun, violations: Violation[]): Promise<void> => {
  const { mode, after, before, base } = run;
  const gridIssues = tableGridIssues(after.doc);
  if (gridIssues.length > 0 && tableGridIssues(before.doc).length === 0) {
    violations.push({ kind: "table-grid", mode, detail: gridIssues.join("; ") });
  }
  let saved: Awaited<ReturnType<typeof saveHarnessState>> | null = null;
  try {
    saved = await saveHarnessState(after, base);
  } catch (error) {
    violations.push({ kind: "invalid-model", mode, detail: errorText(error) });
  }
  if (saved) {
    try {
      const back = await readBack(saved.bytes);
      const blocks = compareSummaries(summarizeState(after), back.summary);
      if (blocks) {
        violations.push({ kind: "readback-blocks", mode, detail: blocks });
      }
      const painted = compareSummaries(summarizeEffectiveParagraphs(after), back.effective);
      if (painted) {
        violations.push({ kind: "readback-painted", mode, detail: painted });
      }
      const markdown = compareText(modelMarkdown(saved.model), back.markdown);
      if (markdown) {
        violations.push({ kind: "readback-markdown", mode, detail: markdown });
      }
    } catch (error) {
      violations.push({
        kind: "readback-blocks",
        mode,
        detail: `reopen threw ${errorText(error)}`,
      });
    }
  }

  try {
    const undone = exhaustHistory(after, undo);
    if (!undone.doc.eq(before.doc)) {
      violations.push({
        kind: "undo",
        mode,
        detail:
          compareSummaries(summarizeState(before), summarizeState(undone)) ??
          describeDocDifference(before.doc, undone.doc),
      });
    } else {
      const redone = exhaustHistory(undone, redo);
      if (!redone.doc.eq(after.doc)) {
        violations.push({
          kind: "redo",
          mode,
          detail:
            compareSummaries(summarizeState(after), summarizeState(redone)) ??
            describeDocDifference(after.doc, redone.doc),
        });
      }
    }
  } catch (error) {
    violations.push({ kind: "undo", mode, detail: `undo threw ${errorText(error)}` });
  }
};

export type CaseResult = {
  /** Outcome per mode; absent when the placement has no meaning in this shape. */
  runs: Partial<Record<EditorMode, ModeRun["status"]>>;
  violations: Violation[];
};

export const runConformanceCase = async (
  shape: DocumentShape,
  operation: ConformanceOperation,
  placement: SelectionPlacement,
): Promise<CaseResult | null> => {
  const loaded = await loadShape(shape);
  const runs: Partial<Record<EditorMode, ModeRun>> = {};
  for (const mode of EDITOR_MODES) {
    const run = runMode(loaded, shape, operation, placement, mode);
    if (!run) {
      return null;
    }
    runs[mode] = run;
  }
  const violations: Violation[] = [];
  const editing = runs.editing;
  const suggesting = runs.suggesting;
  if (!editing || !suggesting) {
    return null;
  }

  for (const run of [editing, suggesting]) {
    if (run.status === "threw") {
      violations.push({ kind: "threw", mode: run.mode, detail: errorText(run.error) });
    } else if (run.status === "changed") {
      // oxlint-disable-next-line no-await-in-loop -- modes are checked one after the other
      await checkChangedState(run, violations);
    }
  }

  // A refusal is `false`; an operation that reports nothing (typing, paste, a
  // claimed key) and leaves the document as it was is refusing just as silently.
  if (
    editing.status === "changed" &&
    (suggesting.status === "refused" || suggesting.status === "unchanged")
  ) {
    violations.push({
      kind: "silent-refusal",
      mode: "suggesting",
      detail: `applies in editing mode but ${suggesting.status === "refused" ? "is refused" : "changes nothing"} in suggesting mode, without a reason`,
    });
  }

  if (suggesting.status === "changed") {
    try {
      const rejectedOriginal = observe(
        resolveAllChanges(suggesting.before, "reject"),
        suggesting.base,
      );
      const rejected = observe(resolveAllChanges(suggesting.after, "reject"), suggesting.base);
      const rejectDifference = compareObservations(rejectedOriginal, rejected);
      if (rejectDifference && operation.suggesting !== "direct") {
        violations.push({ kind: "reject-mismatch", mode: "suggesting", detail: rejectDifference });
      }
      if (editing.status === "changed") {
        const acceptedEditing = observe(resolveAllChanges(editing.after, "accept"), editing.base);
        const accepted = observe(resolveAllChanges(suggesting.after, "accept"), suggesting.base);
        const acceptDifference = compareObservations(acceptedEditing, accepted);
        if (acceptDifference) {
          violations.push({
            kind: "accept-mismatch",
            mode: "suggesting",
            detail: acceptDifference,
          });
        }
      }
      if (operation.suggesting !== "direct") {
        const reopenedOriginal = await observeReopened(
          resolveAllChanges(suggesting.before, "reject"),
          suggesting.base,
        );
        const reopenedRejected = await observeReopened(
          resolveAllChanges(suggesting.after, "reject"),
          suggesting.base,
        );
        const reopenedRejectDifference = compareObservations(reopenedOriginal, reopenedRejected);
        if (reopenedRejectDifference && !rejectDifference) {
          violations.push({
            kind: "reject-mismatch",
            mode: "suggesting",
            detail: `after reopen: ${reopenedRejectDifference}`,
          });
        }
        if (editing.status === "changed") {
          const reopenedEditing = await observeReopened(
            resolveAllChanges(editing.after, "accept"),
            editing.base,
          );
          const reopenedAccepted = await observeReopened(
            resolveAllChanges(suggesting.after, "accept"),
            suggesting.base,
          );
          const reopenedAcceptDifference = compareObservations(reopenedEditing, reopenedAccepted);
          if (reopenedAcceptDifference) {
            violations.push({
              kind: "accept-mismatch",
              mode: "suggesting",
              detail: `after reopen: ${reopenedAcceptDifference}`,
            });
          }
        }
      }
    } catch (error) {
      violations.push({
        kind: "reject-mismatch",
        mode: "suggesting",
        detail: `resolving changes threw ${errorText(error)}`,
      });
    }
  }

  return {
    runs: { editing: editing.status, suggesting: suggesting.status },
    violations,
  };
};
