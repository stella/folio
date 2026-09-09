/**
 * List Extension — list commands + keymaps
 *
 * No schema contribution — lists use paragraph attrs (numPr).
 * Provides: toggle bullet/number, indent/outdent, enter/backspace handling.
 */

import { panic } from "better-result";
import type { Command, EditorState } from "prosemirror-state";

import { expectParagraphAttrs } from "../../attrs";
import {
  hasSerializableParagraphPropertyChange,
  PPR_CHANGE_SCOPED_ATTR_KEYS,
} from "../../commands/propertyChangeScope";
import { makeRevisionInfo, SUGGESTION_META } from "../../plugins/suggestionMode";
import { CLEARED_LIST_RENDERING_ATTRS, LIST_RENDERING_ATTR_KEYS } from "../../listMarker";
import { getDocumentNumbering } from "../../plugins/documentNumbering";
import { listLevelAttrPatch } from "../../styles/resolvedStyleAttrs";
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

function appendParagraphPropertyChange(
  attrs: Record<string, unknown>,
  existing: ParagraphPropertyChangeAttrs[] | undefined,
  previousFormatting: Record<string, unknown>,
  rev: { id: number; author: string; date: string },
): Record<string, unknown> {
  return {
    ...attrs,
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

function getPreviousListFormatting(attrs: Record<string, unknown>): Record<string, unknown> {
  const previousFormatting: Record<string, unknown> = {};
  // Rejecting a pPrChange restores the stored record WHOLESALE within the
  // CT_PPrBase scope (a scoped key absent from the record resets to null —
  // see propertyChangeScope.ts). Snapshot every non-null in-scope attr so a
  // reject cannot wipe formatting the list toggle never touched.
  for (const key of PPR_CHANGE_SCOPED_ATTR_KEYS) {
    const value = attrs[key];
    if (value != null) {
      previousFormatting[key] = value;
    }
  }
  // List-rendering bookkeeping snapshots with explicit nulls: these attrs are
  // outside the wholesale scope, so only recorded keys restore on reject.
  previousFormatting["numPr"] = attrs["numPr"] ?? null;
  for (const key of LIST_RENDERING_ATTR_KEYS) {
    previousFormatting[key] = attrs[key] ?? null;
  }
  return previousFormatting;
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

        let nextAttrs: Record<string, unknown>;

        if (isInSameList) {
          nextAttrs = {
            ...node.attrs,
            numPr: null,
            ...CLEARED_LIST_RENDERING_ATTRS,
          };
        } else {
          const isBullet = numId === 1;
          nextAttrs = {
            ...node.attrs,
            ...CLEARED_LIST_RENDERING_ATTRS,
            numPr: { numId, ilvl: node.attrs["numPr"]?.ilvl || 0 },
            listIsBullet: isBullet,
            listNumFmt: isBullet ? null : "decimal",
          };
        }

        if (rev) {
          const existing = expectParagraphAttrs(node)._propertyChanges;
          nextAttrs = appendParagraphPropertyChange(
            nextAttrs,
            existing,
            getPreviousListFormatting(node.attrs),
            rev,
          );
        }

        tr = tr.setNodeMarkup(pos, undefined, nextAttrs);
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

const attrsForListLevel = (
  state: EditorState,
  attrs: ParagraphAttrs,
  level: number,
): Record<string, unknown> => {
  const numPr = attrs.numPr;
  if (numPr?.numId === undefined) {
    panic("Cannot change the level of a list without a numbering id");
  }
  return {
    ...attrs,
    ...listLevelAttrPatch(attrs, { numId: numPr.numId, ilvl: level }, getDocumentNumbering(state)),
  };
};

const increaseListLevel: Command = (state, dispatch) => {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== "paragraph") {
    return false;
  }
  if (!paragraph.attrs["numPr"]) {
    return false;
  }

  const currentLevel = paragraph.attrs["numPr"].ilvl || 0;
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
        ...attrsForListLevel(state, expectParagraphAttrs(paragraph), currentLevel + 1),
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
  if (!paragraph.attrs["numPr"]) {
    return false;
  }

  const currentLevel = paragraph.attrs["numPr"].ilvl || 0;

  if (!dispatch) {
    return true;
  }

  const paragraphPos = $from.before($from.depth);

  if (currentLevel <= 0) {
    dispatch(
      state.tr
        .setNodeMarkup(paragraphPos, undefined, {
          ...paragraph.attrs,
          numPr: null,
          ...CLEARED_LIST_RENDERING_ATTRS,
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
          ...attrsForListLevel(state, expectParagraphAttrs(paragraph), currentLevel - 1),
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
    if (node.type.name === "paragraph" && node.attrs["numPr"] && !seen.has(pos)) {
      seen.add(pos);
      tr = tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        numPr: null,
        ...CLEARED_LIST_RENDERING_ATTRS,
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
  return !!paragraph.attrs["numPr"]?.numId;
}

export function getListInfo(state: EditorState): { numId: number; ilvl: number } | null {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== "paragraph") {
    return null;
  }
  if (!paragraph.attrs["numPr"]?.numId) {
    return null;
  }

  return {
    numId: paragraph.attrs["numPr"].numId,
    ilvl: paragraph.attrs["numPr"].ilvl || 0,
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

    const numPr = paragraph.attrs["numPr"];
    if (!numPr) {
      return false;
    }

    if (paragraph.textContent.length > 0) {
      return false;
    }

    if (dispatch) {
      const tr = state.tr.setNodeMarkup($from.before(), undefined, {
        ...paragraph.attrs,
        numPr: null,
        ...CLEARED_LIST_RENDERING_ATTRS,
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

    const numPr = paragraph.attrs["numPr"];
    if (!numPr) {
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

    const numPr = paragraph.attrs["numPr"];
    if (!numPr) {
      return false;
    }

    if (dispatch) {
      const tr = state.tr.setNodeMarkup($from.before(), undefined, {
        ...paragraph.attrs,
        numPr: null,
        ...CLEARED_LIST_RENDERING_ATTRS,
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
      if (node.type.name === "paragraph" && node.attrs["numPr"]) {
        const attrs = expectParagraphAttrs(node);
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
        tr = tr.setNodeMarkup(
          pos,
          undefined,
          attrsForListLevel(state, attrs, (attrs.numPr?.ilvl ?? 0) + 1),
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
    const positions: { pos: number; attrs: ParagraphAttrs }[] = [];
    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === "paragraph" && node.attrs["numPr"]) {
        positions.push({ pos, attrs: expectParagraphAttrs(node) });
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
          tr = tr.setNodeMarkup(pos, undefined, {
            ...attrs,
            numPr: null,
            ...CLEARED_LIST_RENDERING_ATTRS,
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
