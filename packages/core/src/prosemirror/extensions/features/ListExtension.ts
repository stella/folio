/**
 * List Extension — list commands + keymaps
 *
 * No schema contribution — lists use paragraph attrs (numPr).
 * Provides: toggle bullet/number, indent/outdent, enter/backspace handling.
 */

import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";

import { expectParagraphAttrs } from "../../attrs";
import {
  hasSerializableParagraphPropertyChange,
  paragraphPropertiesSnapshot,
} from "../../commands/propertyChangeScope";
import { attrsWithParagraphIndentationTransition } from "../../paragraphIndentation";
import {
  applyParagraphPropertyProjection,
  patchParagraphPropertyProjection,
  splitParagraphWithProperties,
  type ParagraphPropertyProjection,
} from "../../paragraphPropertyMutation";
import { makeRevisionInfo, SUGGESTION_META } from "../../plugins/suggestionMode";
import { LIST_RENDERING_ATTR_KEYS } from "../../listMarker";
import { getDocumentNumbering } from "../../plugins/documentNumbering";
import {
  paragraphPropertiesForListLevelTransition,
  paragraphPropertiesForListRemoval,
} from "../../styles/resolvedStyleAttrs";
import { createExtension } from "../create";
import { goToNextCell, goToPrevCell } from "../nodes/TableExtension";
import { Priority } from "../types";
import type { ExtensionRuntime } from "../types";
import type { ParagraphAttrs, ParagraphPropertyChangeAttrs } from "../../schema/nodes";

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

// ============================================================================
// TRACKED PARAGRAPH-PROPERTY CHANGE (suggesting mode)
// ============================================================================

function paragraphPropertyChangePatch(
  existing: ParagraphPropertyChangeAttrs[] | undefined,
  previousFormatting: Record<string, unknown>,
  rev: { id: number; author: string; date: string },
): Record<string, unknown> {
  return {
    _propertyChanges: [
      ...(existing ?? []),
      {
        type: "paragraphPropertyChange",
        info: { id: rev.id, author: rev.author, date: rev.date },
        previousFormatting,
      },
    ],
  };
}

function getPreviousListFormatting(node: PMNode): Record<string, unknown> {
  const previousFormatting: Record<string, unknown> = paragraphPropertiesSnapshot(node);
  const attrs = node.attrs;
  // List-rendering bookkeeping snapshots with explicit nulls: these attrs are
  // outside the wholesale scope, so only recorded keys restore on reject.
  for (const key of LIST_RENDERING_ATTR_KEYS) {
    previousFormatting[key] = attrs[key] ?? null;
  }
  return previousFormatting;
}

type ActiveListParagraphAttrs = ParagraphAttrs & {
  numPr: NonNullable<ParagraphAttrs["numPr"]> & { numId: number };
};

function hasActiveListNumbering(attrs: ParagraphAttrs): attrs is ActiveListParagraphAttrs {
  return attrs.numPr?.numId !== undefined && attrs.numPr.numId !== 0;
}

// ============================================================================
// LIST COMMANDS
// ============================================================================

function toggleList(numId: number): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    const paragraph = $from.parent;
    if (paragraph.type.name !== "paragraph") {
      return false;
    }

    const currentNumPr = paragraph.attrs["numPr"];
    const isInSameList = currentNumPr?.numId === numId;

    const rev = makeRevisionInfo(state);
    if (rev) {
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

    let tr = state.tr;
    const seen = new Set<number>();

    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === "paragraph" && !seen.has(pos)) {
        seen.add(pos);

        const attrs = expectParagraphAttrs(node);
        let projection = isInSameList
          ? paragraphPropertiesForListRemoval(attrs)
          : paragraphPropertiesForListLevelTransition({
              attrs,
              numPr: { numId, ilvl: attrs.numPr?.ilvl ?? 0 },
              numbering: getDocumentNumbering(state),
            });

        if (rev) {
          projection = patchParagraphPropertyProjection({
            projection,
            patch: paragraphPropertyChangePatch(
              attrs._propertyChanges,
              getPreviousListFormatting(node),
              rev,
            ),
          });
        }

        applyParagraphPropertyProjection({
          transaction: tr,
          pos,
          projection,
          source: { type: "preserve" },
        });
      }
    });

    if (rev) {
      tr.setMeta(SUGGESTION_META, true);
    }

    dispatch(tr.scrollIntoView());
    return true;
  };
}

export const toggleBulletList: Command = (state, dispatch) => toggleList(1)(state, dispatch);

export const toggleNumberedList: Command = (state, dispatch) => toggleList(2)(state, dispatch);

const propertiesForListLevel = (
  state: EditorState,
  attrs: ParagraphAttrs,
  level: number,
): ParagraphPropertyProjection => {
  if (!hasActiveListNumbering(attrs)) {
    panic("Cannot change the level of a list without a numbering id");
  }
  return paragraphPropertiesForListLevelTransition({
    attrs,
    numPr: { numId: attrs.numPr.numId, ilvl: level },
    numbering: getDocumentNumbering(state),
  });
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

  const currentLevel = attrs.numPr.ilvl || 0;
  if (currentLevel >= 8) {
    return false;
  }

  if (!dispatch) {
    return true;
  }

  const paragraphPos = $from.before($from.depth);

  const tr = state.tr;
  applyParagraphPropertyProjection({
    transaction: tr,
    pos: paragraphPos,
    projection: propertiesForListLevel(state, attrs, currentLevel + 1),
    source: { type: "preserve" },
  });
  dispatch(tr.scrollIntoView());

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

  const currentLevel = attrs.numPr.ilvl || 0;

  if (!dispatch) {
    return true;
  }

  const paragraphPos = $from.before($from.depth);

  if (currentLevel <= 0) {
    const tr = state.tr;
    applyParagraphPropertyProjection({
      transaction: tr,
      pos: paragraphPos,
      projection: paragraphPropertiesForListRemoval(attrs),
      source: { type: "preserve" },
    });
    const cleared = tr.doc.nodeAt(paragraphPos);
    if (cleared?.type.name !== "paragraph") {
      panic("Clearing list properties lost the paragraph");
    }
    const transition = attrsWithParagraphIndentationTransition(expectParagraphAttrs(cleared), {
      type: "force-visible-zero",
      sides: ["left", "firstLine"],
    });
    if (transition.type === "unsupported") {
      panic("Clearing paragraph numbering left an active numbering-indent owner.");
    }
    applyParagraphPropertyProjection({
      projection: transition.projection,
      source: { type: "preserve" },
      pos: paragraphPos,
      transaction: tr,
    });
    dispatch(tr.scrollIntoView());
  } else {
    const tr = state.tr;
    applyParagraphPropertyProjection({
      transaction: tr,
      pos: paragraphPos,
      projection: propertiesForListLevel(state, attrs, currentLevel - 1),
      source: { type: "preserve" },
    });
    dispatch(tr.scrollIntoView());
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
      applyParagraphPropertyProjection({
        transaction: tr,
        pos,
        projection: paragraphPropertiesForListRemoval(expectParagraphAttrs(node)),
        source: { type: "preserve" },
      });
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
    ilvl: attrs.numPr.ilvl || 0,
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
      const tr = state.tr;
      applyParagraphPropertyProjection({
        transaction: tr,
        pos: $from.before(),
        projection: paragraphPropertiesForListRemoval(attrs),
        source: { type: "preserve" },
      });
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

      splitParagraphWithProperties({
        transaction: tr,
        pos,
        transition: { type: "split-left-created-right-retains" },
        typesAfter: [
          {
            type: state.schema.nodes["paragraph"]!,
            attrs: { ...paragraph.attrs },
          },
        ],
      });

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
      const tr = state.tr;
      applyParagraphPropertyProjection({
        transaction: tr,
        pos: $from.before(),
        projection: paragraphPropertiesForListRemoval(attrs),
        source: { type: "preserve" },
      });
      dispatch(tr);
    }
    return true;
  };
}

function increaseListIndent(): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    // Collect all list paragraphs in the selection range
    const positions: { pos: number; attrs: ParagraphAttrs }[] = [];
    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === "paragraph") {
        const attrs = expectParagraphAttrs(node);
        if (!hasActiveListNumbering(attrs)) {
          return;
        }
        const currentLevel = attrs.numPr?.ilvl ?? 0;
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
        applyParagraphPropertyProjection({
          transaction: tr,
          pos,
          projection: propertiesForListLevel(state, attrs, (attrs.numPr?.ilvl ?? 0) + 1),
          source: { type: "preserve" },
        });
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
    const positions: { pos: number; attrs: ParagraphAttrs }[] = [];
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
        const currentLevel = attrs.numPr?.ilvl ?? 0;
        if (currentLevel <= 0) {
          applyParagraphPropertyProjection({
            transaction: tr,
            pos,
            projection: paragraphPropertiesForListRemoval(attrs),
            source: { type: "preserve" },
          });
          const cleared = tr.doc.nodeAt(pos);
          if (cleared?.type.name !== "paragraph") {
            panic("Clearing list properties lost the paragraph");
          }
          const transition = attrsWithParagraphIndentationTransition(expectParagraphAttrs(cleared), {
            type: "force-visible-zero",
            sides: ["left", "firstLine"],
          });
          if (transition.type === "unsupported") {
            panic("Clearing paragraph numbering left an active numbering-indent owner.");
          }
          applyParagraphPropertyProjection({
            projection: transition.projection,
            source: { type: "preserve" },
            pos,
            transaction: tr,
          });
        } else {
          applyParagraphPropertyProjection({
            transaction: tr,
            pos,
            projection: propertiesForListLevel(state, attrs, currentLevel - 1),
            source: { type: "preserve" },
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
// EXTENSION
// ============================================================================

export const ListExtension = createExtension({
  name: "list",
  priority: Priority.High, // Must be before base keymap
  onSchemaReady(): ExtensionRuntime {
    return {
      commands: {
        toggleBulletList: () => toggleBulletList,
        toggleNumberedList: () => toggleNumberedList,
        increaseListLevel: () => increaseListLevel,
        decreaseListLevel: () => decreaseListLevel,
        removeList: () => removeList,
      },
      keyboardShortcuts: {
        Tab: chainCommands(goToNextCell(), increaseListIndent(), insertTab()),
        "Shift-Tab": chainCommands(goToPrevCell(), decreaseListIndent()),
        "Shift-Enter": () => false, // Let base keymap handle this
        Enter: chainCommands(exitListOnEmptyEnter(), splitListItem()),
        Backspace: backspaceExitList(),
      },
    };
  },
});
