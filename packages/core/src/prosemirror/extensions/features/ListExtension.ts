/**
 * List Extension — list commands + keymaps
 *
 * No schema contribution — lists use paragraph attrs (numPr).
 * Provides: toggle bullet/number, indent/outdent, enter/backspace handling.
 */

import { panic } from "better-result";
import { InputRule, inputRules, undoInputRule } from "prosemirror-inputrules";
import type { Node as PMNode } from "prosemirror-model";
import { Plugin, type Command, type EditorState, type Transaction } from "prosemirror-state";

import { expectParagraphAttrs } from "../../attrs";
import {
  paragraphNumberingLevel,
  paragraphNumberingReferenceId,
} from "../../../docx/numberingReference";
import { LIST_MARKER_PATTERN, listRequestsForMarker } from "../../listAutoformatMarkers";
import { CLEARED_LIST_RENDERING_ATTRS } from "../../listMarker";
import {
  applyParagraphUpdates,
  continueNumbering,
  listItemAttrs,
  resolveListTarget,
  type ListRequest,
  restartNumbering,
  setNumberingValue,
} from "../../listNumbering";
import { resolveListState, type ListType } from "../../listState";
import { removedNumberingAttr, type ParagraphNumberingAttr } from "../../numberingAttr";
import { getDocumentNumbering } from "../../plugins/documentNumbering";
import {
  makeRevisionInfo,
  SUGGESTED_TEXT_INPUT_META,
  suggestRangeDeletion,
} from "../../plugins/suggestionMode";
import { listLevelAttrPatch } from "../../styles/resolvedStyleAttrs";
import { createExtension } from "../create";
import { goToNextCell, goToPrevCell } from "../nodes/TableExtension";
import { Priority } from "../types";
import type { ExtensionRuntime } from "../types";
import type { ParagraphAttrs, ParagraphAttrsPatch } from "../../schema/nodes";

// ============================================================================
// CHAIN COMMANDS HELPER
// ============================================================================

function chainCommands(...commands: Command[]): Command {
  return (state, dispatch, view) => {
    for (const cmd of commands) {
      if (cmd(state, dispatch, view)) {
        return true;
      }
    }
    return false;
  };
}

function clearListAttrs(attrs: ParagraphAttrs): Record<string, unknown> {
  return {
    ...attrs,
    numPr: removedNumberingAttr(attrs.numPrFromStyle),
    ...CLEARED_LIST_RENDERING_ATTRS,
  };
}

type ActiveListParagraphAttrs = ParagraphAttrs & {
  numPr: Extract<ParagraphNumberingAttr, { kind: "reference" }>;
};

function hasActiveListNumbering(attrs: ParagraphAttrs): attrs is ActiveListParagraphAttrs {
  return attrs.numPr?.kind === "reference";
}

// ============================================================================
// LIST COMMANDS
// ============================================================================

type ActiveListType = Exclude<ListType, "none">;

type ParagraphInRange = { pos: number; node: PMNode };

const paragraphsBetween = (doc: PMNode, from: number, to: number): ParagraphInRange[] => {
  const paragraphs: ParagraphInRange[] = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    paragraphs.push({ pos, node });
    return false;
  });
  return paragraphs;
};

type NumberParagraphsOptions = {
  state: EditorState;
  /** Positions `from` and `to` are in `tr.doc`. */
  tr: Transaction;
  from: number;
  to: number;
  requests: readonly ListRequest[];
};

/**
 * Make every paragraph from `from` to `to` an item of one list: the list next
 * to them when it matches, else a new one (see `listNumbering.ts`). One
 * target for the whole range, so a selection of plain paragraphs becomes one
 * list rather than one list apiece.
 */
const numberParagraphs = ({ state, tr, from, to, requests }: NumberParagraphsOptions): boolean => {
  const target = resolveListTarget({
    numbering: getDocumentNumbering(state),
    $from: tr.doc.resolve(from),
    $to: tr.doc.resolve(to),
    requests,
  });
  if (!target) {
    return false;
  }
  applyParagraphUpdates({
    tr,
    state,
    updates: paragraphsBetween(tr.doc, from, to).map(({ pos, node }) => ({
      pos,
      node,
      // The definition, not the id, says what the paragraph now renders: the
      // full attr group, so the painter, the next command and the save all
      // read the same level.
      next: listItemAttrs(
        node.attrs,
        {
          numId: target.numId,
          ilvl: paragraphNumberingLevel(expectParagraphAttrs(node).numPr) ?? target.ilvl,
        },
        target.numbering,
      ),
    })),
  });
  return true;
};

function toggleList(intent: ActiveListType): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    const paragraph = $from.parent;
    if (paragraph.type.name !== "paragraph") {
      return false;
    }
    if (!dispatch) {
      return true;
    }

    const numbering = getDocumentNumbering(state);
    const isItemOfKind = (node: PMNode): boolean =>
      resolveListState(numbering, expectParagraphAttrs(node).numPr).type === intent;
    const tr = state.tr;
    if (isItemOfKind(paragraph)) {
      // Only the items of that kind leave it: a plain paragraph in the
      // selection has nothing to clear and nothing to track.
      applyParagraphUpdates({
        tr,
        state,
        updates: paragraphsBetween(state.doc, $from.pos, $to.pos)
          .filter(({ node }) => isItemOfKind(node))
          .map(({ pos, node }) => ({
            pos,
            node,
            next: clearListAttrs(expectParagraphAttrs(node)),
          })),
      });
    } else {
      numberParagraphs({ state, tr, from: $from.pos, to: $to.pos, requests: [{ kind: intent }] });
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export const toggleBulletList: Command = (state, dispatch) => toggleList("bullet")(state, dispatch);

export const toggleNumberedList: Command = (state, dispatch) =>
  toggleList("numbered")(state, dispatch);

const attrsForListLevel = (
  state: EditorState,
  attrs: ParagraphAttrs,
  level: number,
): ParagraphAttrsPatch => {
  if (!hasActiveListNumbering(attrs)) {
    panic("Cannot change the level of a list without a numbering id");
  }
  return {
    ...attrs,
    ...listLevelAttrPatch(
      attrs,
      { numId: attrs.numPr.numId, ilvl: level },
      getDocumentNumbering(state),
    ),
  };
};

const increaseListLevel: Command = (state, dispatch) => {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== "paragraph") {
    return false;
  }
  const attrs = expectParagraphAttrs(paragraph);
  if (!hasActiveListNumbering(attrs)) {
    return false;
  }

  const currentLevel = attrs.numPr.ilvl ?? 0;
  if (currentLevel >= 8) {
    return false;
  }

  if (!dispatch) {
    return true;
  }

  const paragraphPos = $from.before($from.depth);

  dispatch(
    state.tr
      .setNodeMarkup(paragraphPos, undefined, {
        ...attrsForListLevel(state, attrs, currentLevel + 1),
      })
      .scrollIntoView(),
  );

  return true;
};

const decreaseListLevel: Command = (state, dispatch) => {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== "paragraph") {
    return false;
  }
  const attrs = expectParagraphAttrs(paragraph);
  if (!hasActiveListNumbering(attrs)) {
    return false;
  }

  const currentLevel = attrs.numPr.ilvl ?? 0;

  if (!dispatch) {
    return true;
  }

  const paragraphPos = $from.before($from.depth);

  if (currentLevel <= 0) {
    dispatch(
      state.tr
        .setNodeMarkup(paragraphPos, undefined, {
          ...clearListAttrs(attrs),
          indentLeft: null,
          indentFirstLine: null,
          hangingIndent: null,
        })
        .scrollIntoView(),
    );
  } else {
    dispatch(
      state.tr
        .setNodeMarkup(paragraphPos, undefined, {
          ...attrsForListLevel(state, attrs, currentLevel - 1),
        })
        .scrollIntoView(),
    );
  }

  return true;
};

const removeList: Command = (state, dispatch) => {
  const { $from, $to } = state.selection;

  if (!dispatch) {
    return true;
  }

  let tr = state.tr;
  const seen = new Set<number>();

  state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
    if (
      node.type.name === "paragraph" &&
      hasActiveListNumbering(expectParagraphAttrs(node)) &&
      !seen.has(pos)
    ) {
      seen.add(pos);
      tr = tr.setNodeMarkup(pos, undefined, clearListAttrs(expectParagraphAttrs(node)));
    }
  });

  dispatch(tr.scrollIntoView());
  return true;
};

// ============================================================================
// LIST QUERY HELPERS (exported for toolbar)
// ============================================================================

export function isInList(state: EditorState): boolean {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== "paragraph") {
    return false;
  }
  return hasActiveListNumbering(expectParagraphAttrs(paragraph));
}

export function getListInfo(state: EditorState): { numId: number; ilvl: number } | null {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== "paragraph") {
    return null;
  }
  const attrs = expectParagraphAttrs(paragraph);
  if (!hasActiveListNumbering(attrs)) {
    return null;
  }

  return {
    numId: attrs.numPr.numId,
    ilvl: attrs.numPr.ilvl ?? 0,
  };
}

// ============================================================================
// KEYMAP COMMANDS
// ============================================================================

function exitListOnEmptyEnter(): Command {
  return (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty) {
      return false;
    }

    const paragraph = $from.parent;
    if (paragraph.type.name !== "paragraph") {
      return false;
    }

    const attrs = expectParagraphAttrs(paragraph);
    if (!hasActiveListNumbering(attrs)) {
      return false;
    }

    if (paragraph.textContent.length > 0) {
      return false;
    }

    if (dispatch) {
      const tr = state.tr.setNodeMarkup($from.before(), undefined, clearListAttrs(attrs));
      dispatch(tr);
    }
    return true;
  };
}

function splitListItem(): Command {
  return (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty) {
      return false;
    }

    const paragraph = $from.parent;
    if (paragraph.type.name !== "paragraph") {
      return false;
    }

    const attrs = expectParagraphAttrs(paragraph);
    if (!hasActiveListNumbering(attrs)) {
      return false;
    }

    if (dispatch) {
      const { tr } = state;
      const pos = $from.pos;

      tr.split(pos, 1, [
        {
          type: state.schema.nodes["paragraph"]!,
          attrs: { ...paragraph.attrs },
        },
      ]);

      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

function backspaceExitList(): Command {
  return (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty) {
      return false;
    }

    if ($from.parentOffset !== 0) {
      return false;
    }

    const paragraph = $from.parent;
    if (paragraph.type.name !== "paragraph") {
      return false;
    }

    const attrs = expectParagraphAttrs(paragraph);
    if (!hasActiveListNumbering(attrs)) {
      return false;
    }

    if (dispatch) {
      const tr = state.tr.setNodeMarkup($from.before(), undefined, clearListAttrs(attrs));
      dispatch(tr);
    }
    return true;
  };
}

function increaseListIndent(): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    // Collect all list paragraphs in the selection range
    const positions: { pos: number; attrs: ActiveListParagraphAttrs }[] = [];
    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === "paragraph") {
        const attrs = expectParagraphAttrs(node);
        if (!hasActiveListNumbering(attrs)) {
          return;
        }
        const currentLevel = attrs.numPr.ilvl ?? 0;
        if (currentLevel < 8) {
          positions.push({ pos, attrs });
        }
      }
    });

    if (positions.length === 0) {
      return false;
    }

    if (dispatch) {
      let tr = state.tr;
      for (const { pos, attrs } of positions) {
        tr = tr.setNodeMarkup(
          pos,
          undefined,
          attrsForListLevel(state, attrs, (attrs.numPr.ilvl ?? 0) + 1),
        );
      }
      dispatch(tr);
    }
    return true;
  };
}

function decreaseListIndent(): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    // Collect all list paragraphs in the selection range
    const positions: { pos: number; attrs: ActiveListParagraphAttrs }[] = [];
    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === "paragraph") {
        const attrs = expectParagraphAttrs(node);
        if (hasActiveListNumbering(attrs)) {
          positions.push({ pos, attrs });
        }
      }
    });

    if (positions.length === 0) {
      return false;
    }

    if (dispatch) {
      let tr = state.tr;
      for (const { pos, attrs } of positions) {
        const currentLevel = attrs.numPr.ilvl ?? 0;
        if (currentLevel <= 0) {
          tr = tr.setNodeMarkup(pos, undefined, {
            ...clearListAttrs(attrs),
            indentLeft: null,
            indentFirstLine: null,
            hangingIndent: null,
          });
        } else {
          tr = tr.setNodeMarkup(pos, undefined, {
            ...attrsForListLevel(state, attrs, currentLevel - 1),
          });
        }
      }
      dispatch(tr);
    }
    return true;
  };
}

function insertTab(): Command {
  return (state, dispatch) => {
    const { schema } = state;
    const tabType = schema.nodes["tab"];

    if (!tabType) {
      return false;
    }

    if (dispatch) {
      const tr = state.tr.replaceSelectionWith(tabType.create());
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

// goToNextCell/goToPrevCell are imported at the top from table extension for chaining

// ============================================================================
// AUTOFORMAT (list markers typed at the start of a paragraph)
// ============================================================================

/** Whether a typed marker may become a list here at all. */
const acceptsListMarker = (paragraph: PMNode): boolean =>
  paragraph.type.name === "paragraph" &&
  // Toggling a list that already carries numbering would change or remove it.
  paragraphNumberingReferenceId(expectParagraphAttrs(paragraph).numPr) === undefined;

/**
 * Replace a typed marker with the list the toolbar button produces: the same
 * target resolution and the same attrs, so an autoformatted list saves exactly
 * like a clicked one. The marker's format and value choose the list (see
 * `listAutoformatMarkers.ts`).
 */
const listAutoformatRule = (): InputRule =>
  new InputRule(LIST_MARKER_PATTERN, (state, match, start, end) => {
    const { $from } = state.selection;
    if (!acceptsListMarker($from.parent)) {
      return null;
    }
    // A rule matches a window of text ending at the caret, so `^` alone would
    // also match a marker typed mid-paragraph in a long one.
    if (start !== $from.start()) {
      return null;
    }
    // While suggesting, the suggestion plugin claims the typed text first;
    // `suggestedListAutoformat` converts the marker after it.
    if (makeRevisionInfo(state)) {
      return null;
    }
    const requests = listRequestsForMarker(match[0]);
    if (!requests) {
      return null;
    }
    const tr = state.tr.delete(start, end);
    return numberParagraphs({ state, tr, from: start, to: start, requests }) ? tr : null;
  });

/**
 * List autoformat while suggesting. The suggestion plugin has already
 * recorded the typed space as a tracked insertion; this follows it with the
 * marker's removal (a retraction of the author's own typing, or a tracked
 * deletion of text that was there before) and the list as a tracked
 * paragraph-property change. The follow-up is appended to the typing
 * transaction, so one undo reverts both, and it is stored as the input rule's
 * last conversion, so Backspace puts the marker back as it does while editing.
 */
const suggestedListAutoformat = (autoformat: Plugin): Plugin =>
  new Plugin({
    appendTransaction(transactions, _oldState, state) {
      if (!transactions.some((tr) => tr.getMeta(SUGGESTED_TEXT_INPUT_META) === " ")) {
        return null;
      }
      const { $head, empty } = state.selection;
      if (!empty || !makeRevisionInfo(state) || !acceptsListMarker($head.parent)) {
        return null;
      }
      const requests = listRequestsForMarker(
        $head.parent.textBetween(0, $head.parentOffset, null, "\ufffc"),
      );
      if (!requests) {
        return null;
      }
      const start = $head.start();
      const tr = state.tr;
      if (!suggestRangeDeletion(state, tr, start, $head.pos)) {
        return null;
      }
      const from = tr.mapping.map(start);
      if (!numberParagraphs({ state, tr, from, to: from, requests })) {
        return null;
      }
      const caret = tr.mapping.map($head.pos);
      tr.setMeta(autoformat, { transform: tr, from: caret, to: caret, text: "" });
      return tr;
    },
  });

const listAutoformatPlugins = (): Plugin[] => {
  const autoformat = inputRules({ rules: [listAutoformatRule()] });
  return [autoformat, suggestedListAutoformat(autoformat)];
};

// ============================================================================
// EXTENSION
// ============================================================================

export const ListExtension = createExtension({
  name: "list",
  priority: Priority.High, // Must be before base keymap
  onSchemaReady(): ExtensionRuntime {
    return {
      plugins: listAutoformatPlugins(),
      commands: {
        toggleBulletList: () => toggleBulletList,
        toggleNumberedList: () => toggleNumberedList,
        increaseListLevel: () => increaseListLevel,
        decreaseListLevel: () => decreaseListLevel,
        removeList: () => removeList,
        restartNumbering: () => restartNumbering,
        continueNumbering: () => continueNumbering,
        setNumberingValue: (value: number) => setNumberingValue(value),
      },
      keyboardShortcuts: {
        Tab: chainCommands(goToNextCell(), increaseListIndent(), insertTab()),
        "Shift-Tab": chainCommands(goToPrevCell(), decreaseListIndent()),
        "Shift-Enter": () => false, // Let base keymap handle this
        Enter: chainCommands(exitListOnEmptyEnter(), splitListItem()),
        // Backspace right after an autoformat puts the typed marker back.
        Backspace: chainCommands(undoInputRule, backspaceExitList()),
      },
    };
  },
});
