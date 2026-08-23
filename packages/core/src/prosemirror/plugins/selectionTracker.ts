/**
 * Selection Tracker Plugin
 *
 * Tracks selection changes and emits events for toolbar state updates.
 * Provides the current selection context including:
 * - Text formatting at cursor/selection
 * - Paragraph formatting
 * - Selection range information
 */

import { Plugin, PluginKey } from "prosemirror-state";
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { expectCommentMarkAttrs, expectTrackedChangeMarkAttrs } from "../attrs";
import { extractSelectionSnapshot } from "../selectionState";
import type { TextFormatting, ParagraphFormatting } from "../../types/document";

/**
 * Selection context for toolbar state
 */
export type SelectionContext = {
  /** Whether there's a non-collapsed selection */
  hasSelection: boolean;
  /** Whether selection spans multiple paragraphs */
  isMultiParagraph: boolean;
  /** Current text formatting at cursor/selection */
  textFormatting: TextFormatting;
  /** Current paragraph formatting */
  paragraphFormatting: ParagraphFormatting;
  /** Start paragraph index */
  startParagraphIndex: number;
  /** End paragraph index */
  endParagraphIndex: number;
  /** Whether cursor is in a list */
  inList: boolean;
  /** List type if in list */
  listType?: "bullet" | "numbered";
  /** List level (0-8) */
  listLevel?: number;
  /** Active comment IDs at cursor position */
  activeCommentIds: number[];
  /** Whether cursor is inside a tracked insertion */
  inInsertion: boolean;
  /** Whether cursor is inside a tracked deletion */
  inDeletion: boolean;
};

/**
 * Plugin key for accessing selection tracker state
 */
export const selectionTrackerKey = new PluginKey<SelectionContext>("selectionTracker");

/**
 * Callback type for selection changes
 */
export type SelectionChangeCallback = (context: SelectionContext) => void;

/**
 * Extract selection context from editor state
 */
export function extractSelectionContext(state: EditorState): SelectionContext {
  const { selection } = state;
  const { empty } = selection;
  const snapshot = extractSelectionSnapshot(state);

  const paragraphFormatting: ParagraphFormatting =
    snapshot.styleId === null
      ? snapshot.paragraphFormatting
      : { ...snapshot.paragraphFormatting, styleId: snapshot.styleId };

  // List detection
  const numPr = snapshot.paragraphFormatting.numPr;
  const inList = !!numPr?.numId;
  let listType: "bullet" | "numbered" | undefined;
  if (numPr?.numId === 1) {
    listType = "bullet";
  } else if (numPr?.numId) {
    listType = "numbered";
  }
  const listLevel = numPr?.ilvl;

  // Comment and tracked change detection
  const allMarks = state.storedMarks || (empty ? selection.$from.marks() : []);
  const activeCommentIds: number[] = [];
  let inInsertion = false;
  let inDeletion = false;

  for (const mark of allMarks) {
    if (mark.type.name === "comment") {
      activeCommentIds.push(expectCommentMarkAttrs(mark).commentId);
    }
    if (mark.type.name === "insertion") {
      expectTrackedChangeMarkAttrs(mark);
      inInsertion = true;
    }
    if (mark.type.name === "deletion") {
      expectTrackedChangeMarkAttrs(mark);
      inDeletion = true;
    }
  }

  return {
    hasSelection: snapshot.hasSelection,
    isMultiParagraph: snapshot.isMultiParagraph,
    textFormatting: snapshot.textFormatting,
    paragraphFormatting,
    startParagraphIndex: snapshot.startParagraphIndex,
    endParagraphIndex: snapshot.endParagraphIndex,
    inList,
    ...(listType !== undefined ? { listType } : {}),
    ...(listLevel !== undefined ? { listLevel } : {}),
    activeCommentIds,
    inInsertion,
    inDeletion,
  };
}

/**
 * Create selection tracker plugin
 */
export function createSelectionTrackerPlugin(onSelectionChange?: SelectionChangeCallback): Plugin {
  return new Plugin({
    key: selectionTrackerKey,

    state: {
      init(_, state) {
        return extractSelectionContext(state);
      },

      apply(tr, prevContext, _, newState) {
        // Only recalculate if selection or doc changed
        if (!tr.selectionSet && !tr.docChanged) {
          return prevContext;
        }

        const newContext = extractSelectionContext(newState);

        // Notify callback if context changed
        if (onSelectionChange && !contextsEqual(prevContext, newContext)) {
          // Defer to next tick to avoid dispatch during dispatch
          setTimeout(() => onSelectionChange(newContext), 0);
        }

        return newContext;
      },
    },

    view() {
      return {
        update(view: EditorView, prevState: EditorState) {
          if (!onSelectionChange) {
            return;
          }
          // Only emit on selection/doc changes
          if (view.state.selection.eq(prevState.selection) && view.state.doc.eq(prevState.doc)) {
            return;
          }
          // Reuse context already computed in state.apply() — avoid double doc walk
          const context = selectionTrackerKey.getState(view.state);
          if (context) {
            onSelectionChange(context);
          }
        },
      };
    },
  });
}

function arraysEqual(a: number[] | undefined, b: number[] | undefined): boolean {
  if (a === b) {
    return true;
  }
  if (!a || !b || a.length !== b.length) {
    return false;
  }
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
}

/**
 * Compare two selection contexts for equality
 */
function contextsEqual(a: SelectionContext, b: SelectionContext): boolean {
  return (
    a.hasSelection === b.hasSelection &&
    a.isMultiParagraph === b.isMultiParagraph &&
    a.startParagraphIndex === b.startParagraphIndex &&
    a.endParagraphIndex === b.endParagraphIndex &&
    a.inList === b.inList &&
    a.listType === b.listType &&
    a.listLevel === b.listLevel &&
    a.inInsertion === b.inInsertion &&
    a.inDeletion === b.inDeletion &&
    arraysEqual(a.activeCommentIds, b.activeCommentIds) &&
    JSON.stringify(a.textFormatting) === JSON.stringify(b.textFormatting) &&
    JSON.stringify(a.paragraphFormatting) === JSON.stringify(b.paragraphFormatting)
  );
}

/**
 * Get current selection context from editor state
 */
export function getSelectionContext(state: EditorState): SelectionContext | null {
  return selectionTrackerKey.getState(state) || null;
}
