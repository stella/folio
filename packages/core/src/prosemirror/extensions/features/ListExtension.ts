/**
 * List Extension — list commands + keymaps
 *
 * No schema contribution — lists use paragraph attrs (numPr).
 * Provides: toggle bullet/number, indent/outdent, enter/backspace handling.
 */

import { panic } from "better-result";
import { InputRule, inputRules, undoInputRule } from "prosemirror-inputrules";
import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Plugin, Transaction } from "prosemirror-state";

import { expectParagraphAttrs } from "../../attrs";
import { hasSerializableParagraphPropertyChange } from "../../commands/propertyChangeScope";
import {
  NO_PARAGRAPH_NUMBERING,
  paragraphNumberingLevel,
  paragraphNumberingReferenceId,
} from "../../../docx/numberingReference";
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
import { paragraphNumberingAttr, type ParagraphNumberingAttr } from "../../numberingAttr";
import { getDocumentNumbering } from "../../plugins/documentNumbering";
import { makeRevisionInfo } from "../../plugins/suggestionMode";
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
  // A style supplies the numbering this paragraph is leaving, so the paragraph
  // has to state the cancellation itself — deleting the attr would uncover the
  // style tier and hand the numbering straight back. A cancellation states no
  // level (17.9.18: there is no id left for a level to belong to).
  const numPr =
    paragraphNumberingReferenceId(attrs.numPrFromStyle) === undefined
      ? null
      : paragraphNumberingAttr(NO_PARAGRAPH_NUMBERING);

  return {
    ...attrs,
    numPr,
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

    if (makeRevisionInfo(state)) {
      let hasPendingChange = false;
      state.doc.nodesBetween($from.pos, $to.pos, (node) => {
        if (
          node.type.name === "paragraph" &&
          hasSerializableParagraphPropertyChange(expectParagraphAttrs(node)._propertyChanges)
        ) {
          hasPendingChange = true;
          return false;
        }
        return undefined;
      });
      if (hasPendingChange) {
        return false;
      }
    }

    if (!dispatch) {
      return true;
    }

    const isInSameList =
      resolveListState(getDocumentNumbering(state), expectParagraphAttrs(paragraph).numPr).type ===
      intent;
    const tr = state.tr;
    if (isInSameList) {
      applyParagraphUpdates({
        tr,
        state,
        updates: paragraphsBetween(state.doc, $from.pos, $to.pos).map(({ pos, node }) => ({
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

/**
 * Word's as-you-type markers. Each ends in the space that triggers the rule.
 * `1.` makes a numbered list, like the toolbar button; another typed number
 * stays text rather than being silently renumbered.
 */
const BULLET_AUTOFORMAT = /^[-*] $/u;
const NUMBERED_AUTOFORMAT = /^1\. $/u;

/**
 * Replace a typed marker with the list the toolbar button produces: the same
 * target resolution and the same attrs, so an autoformatted list saves exactly
 * like a clicked one.
 */
const listAutoformatRule = (marker: RegExp, request: ListRequest): InputRule =>
  new InputRule(marker, (state, _match, start, end) => {
    const { $from } = state.selection;
    if ($from.parent.type.name !== "paragraph") {
      return null;
    }
    // A rule matches a window of text ending at the caret, so `^` alone would
    // also match a marker typed mid-paragraph in a long one.
    if (start !== $from.start()) {
      return null;
    }
    // Toggling a list that already carries this numbering would remove it.
    if (paragraphNumberingReferenceId(expectParagraphAttrs($from.parent).numPr) !== undefined) {
      return null;
    }
    // Suggesting mode rewrites typed text as a tracked insertion before any
    // rule runs; autoformatting there would drop that transaction's metadata.
    if (makeRevisionInfo(state)) {
      return null;
    }
    const tr = state.tr.delete(start, end);
    return numberParagraphs({ state, tr, from: start, to: start, requests: [request] }) ? tr : null;
  });

const listAutoformatPlugins = (): Plugin[] => [
  inputRules({
    rules: [
      listAutoformatRule(BULLET_AUTOFORMAT, { kind: "bullet" }),
      listAutoformatRule(NUMBERED_AUTOFORMAT, { kind: "numbered" }),
    ],
  }),
];

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
