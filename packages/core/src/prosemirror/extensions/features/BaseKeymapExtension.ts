/**
 * Base Keymap Extension — wraps prosemirror-commands baseKeymap
 *
 * Priority: Low (150) — must be the last keymap so other extensions can override keys
 */

import {
  baseKeymap,
  splitBlock,
  joinBackward,
  joinForward,
  selectAll,
  selectParentNode,
} from "prosemirror-commands";
import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection, type Command, type Transaction } from "prosemirror-state";
import { canJoin, canSplit } from "prosemirror-transform";

import { expectParagraphAttrs } from "../../attrs";
import { attrsWithParagraphIndentationTransition } from "../../paragraphIndentation";
import {
  applyParagraphPropertyProjection,
  createParagraphNodeFromProjection,
  joinParagraphsWithProperties,
  replaceSelectionThenSplitParagraphWithProperties,
  rebindSplitParagraphProperties,
  splitParagraphWithProperties,
  transitionParagraphProperties,
} from "../../paragraphPropertyMutation";
import { expectParagraphPropertyState } from "../../paragraphPropertyState";
import { recordDeleteParagraphPropertyOwnershipProof } from "../../paragraphPropertyOwnership";
import { getDocumentNumbering } from "../../plugins/documentNumbering";
import { getDocumentStyleResolver } from "../../plugins/documentStyles";
import { inheritedParagraphRunFormatting } from "../../rebaseParagraphRunFormatting";
import { paragraphPropertiesForStyleTransition } from "../../styles/resolvedStyleAttrs";
import type { StyleResolver } from "../../styles/styleResolver";
import {
  tableParagraphStyleContextAtPositions,
  type TableParagraphStyleContext,
} from "../../tableConditionalFormatting";
import { createExtension } from "../create";
import { Priority } from "../types";
import type { ExtensionRuntime, ExtensionContext } from "../types";

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

const deleteSelectionWithParagraphOwnership: Command = (state, dispatch) => {
  if (state.selection.empty) {
    return false;
  }
  if (!dispatch) {
    return true;
  }
  const transaction = state.tr;
  recordDeleteParagraphPropertyOwnershipProof(transaction);
  dispatch(transaction.deleteSelection().scrollIntoView());
  return true;
};

/**
 * Backspace at the start of a paragraph clears first-line indent / hanging indent
 * before joining with the previous paragraph (matches Word behavior).
 */
const clearIndentOnBackspace: Command = (state, dispatch) => {
  const { $cursor } = state.selection as {
    $cursor?: {
      parentOffset: number;
      parent: { type: { name: string }; attrs: Record<string, unknown> };
      pos: number;
      before: () => number;
    };
  };
  if (!$cursor) {
    return false;
  }

  // Only at the very start of a paragraph
  if ($cursor.parentOffset !== 0) {
    return false;
  }
  if ($cursor.parent.type.name !== "paragraph") {
    return false;
  }

  const attrs = $cursor.parent.attrs;
  const hasFirstLine =
    attrs["indentFirstLine"] !== null && (attrs["indentFirstLine"] as number) > 0;
  const hasHanging = !!attrs["hangingIndent"];
  const hasIndentLeft = attrs["indentLeft"] !== null && (attrs["indentLeft"] as number) > 0;

  if (!hasFirstLine && !hasHanging && !hasIndentLeft) {
    return false;
  }

  const transition = attrsWithParagraphIndentationTransition(
    expectParagraphAttrs($cursor.parent as PMNode),
    { type: "force-visible-zero", sides: ["left", "firstLine"] },
  );
  if (transition.type === "unsupported") {
    return false;
  }

  if (dispatch) {
    const pos = $cursor.before();
    const tr = state.tr;
    applyParagraphPropertyProjection({
      projection: transition.projection,
      source: { type: "preserve" },
      pos,
      transaction: tr,
    });
    dispatch(tr.scrollIntoView());
  }
  return true;
};

type ParagraphJoinCandidate = {
  leftContentSize: number;
  joinPosition: number;
};

const paragraphJoinCandidate = (
  state: Parameters<Command>[0],
  direction: "backward" | "forward",
): ParagraphJoinCandidate | null => {
  const { $cursor } = state.selection as {
    $cursor?: {
      before: () => number;
      parent: PMNode;
      parentOffset: number;
    };
  };
  if (!$cursor || $cursor.parent.type.name !== "paragraph") {
    return null;
  }
  const current = $cursor.parent;
  const currentPosition = $cursor.before();
  let left: PMNode | null;
  let right: PMNode | null;
  let leftPosition: number;
  if (direction === "backward") {
    if ($cursor.parentOffset !== 0) {
      return null;
    }
    right = current;
    left = state.doc.resolve(currentPosition).nodeBefore;
    leftPosition = left ? currentPosition - left.nodeSize : -1;
  } else {
    if ($cursor.parentOffset !== current.content.size) {
      return null;
    }
    left = current;
    leftPosition = currentPosition;
    right = state.doc.resolve(currentPosition + current.nodeSize).nodeAfter;
  }
  if (left?.type.name !== "paragraph" || right?.type.name !== "paragraph") {
    return null;
  }
  return {
    leftContentSize: left.content.size,
    joinPosition: leftPosition + left.nodeSize,
  };
};

const withParagraphJoinOwnership =
  (command: Command, direction: "backward" | "forward"): Command =>
  (state, dispatch, view) => {
    if (!dispatch) {
      return command(state, undefined, view);
    }
    const candidate = paragraphJoinCandidate(state, direction);
    if (candidate && canJoin(state.doc, candidate.joinPosition)) {
      const transaction = state.tr;
      const joinedPosition = joinParagraphsWithProperties({
        transaction,
        joinPos: candidate.joinPosition,
        transition: { type: "join-right-paragraph-mark-retains" },
      });
      transaction.setSelection(
        TextSelection.near(transaction.doc.resolve(joinedPosition + candidate.leftContentSize + 1)),
      );
      dispatch(transaction.scrollIntoView());
      return true;
    }
    let captured: Transaction | null = null;
    const handled = command(
      state,
      (transaction) => {
        if (captured !== null) {
          panic("A base join command dispatched more than one transaction");
        }
        captured = transaction;
      },
      view,
    );
    if (!handled) {
      return false;
    }
    if (captured === null) {
      return true;
    }
    dispatch(captured);
    return true;
  };

/**
 * Project the empty paragraph created by Enter from the complete inherited
 * cascade at its table occurrence. A `w:next` style replaces the source
 * style; otherwise the new paragraph keeps the source style while dropping
 * direct paragraph formatting such as borders and indentation.
 */
type ProjectSplitParagraphOptions = {
  context: TableParagraphStyleContext | null | undefined;
  newPara: PMNode;
  resolver: StyleResolver;
  sourcePara: PMNode;
  state: Parameters<Command>[0];
  tr: Transaction;
  useNextStyle: boolean;
};

const projectSplitParagraph = ({
  context,
  resolver,
  sourcePara,
  state,
  tr,
  newPara,
  useNextStyle,
}: ProjectSplitParagraphOptions): Transaction => {
  const sourceStyleId = sourcePara.attrs["styleId"] as string | null | undefined;
  const styleId = useNextStyle
    ? (resolver.getNextStyleId(sourceStyleId) ?? sourceStyleId)
    : sourceStyleId;
  const resolved = resolver.resolveParagraphStyleInTable(styleId, context?.pPr);
  const styleName = styleId ? resolver.getStyle(styleId)?.name : undefined;
  const projection = paragraphPropertiesForStyleTransition({
    attrs: expectParagraphAttrs(newPara),
    identity: { styleId: styleId ?? null, ...(styleName ? { styleName } : {}) },
    numbering: getDocumentNumbering(state),
    resolved,
    styleResolver: resolver,
    ...(context?.rPr ? { tableRunFormatting: context.rPr } : {}),
    transition: { type: "replace-style" },
  });
  const projected = createParagraphNodeFromProjection({
    type: newPara.type,
    projection,
    content: newPara.content,
    marks: newPara.marks,
  });
  const inherited = inheritedParagraphRunFormatting({
    paragraph: projected,
    styleResolver: resolver,
    tableRunFormatting: context?.rPr ?? null,
  });
  const { $from } = tr.selection;
  applyParagraphPropertyProjection({
    transaction: tr,
    pos: $from.before(),
    projection,
    source: { type: "preserve" },
  });
  tr.setStoredMarks(inherited.marks);
  return tr;
};

export const splitBlockClearBorders: Command = (state, dispatch, view) => {
  // Capture source paragraph info BEFORE split (splitBlock resets everything)
  const { $from: preSplitFrom } = state.selection;
  const sourcePara = preSplitFrom.parent.type.name === "paragraph" ? preSplitFrom.parent : null;
  const sourcePropertyState = sourcePara
    ? expectParagraphPropertyState(expectParagraphAttrs(sourcePara)._paragraphPropertyState)
    : null;

  const sourcePosition = sourcePara ? preSplitFrom.before() : null;
  const resolver = getDocumentStyleResolver(state);
  const sourceTableContext =
    resolver && sourcePosition !== null
      ? tableParagraphStyleContextAtPositions(state.doc, resolver).get(sourcePosition)
      : undefined;

  // Intercept splitBlock's transaction so we can modify it before dispatch.
  // This ensures attrs + stored marks are set in a single transaction,
  // avoiding a flash where the empty paragraph has no formatting.
  const splitResult = { ownershipRebound: false, tr: null as Transaction | null };
  const capturingDispatch = dispatch
    ? (tr: Transaction) => {
        splitResult.tr = tr;
      }
    : undefined;

  if (dispatch && sourcePara && sourcePropertyState && sourcePosition !== null) {
    const transaction = state.tr;
    const preflight = state.selection.empty ? transaction : state.tr.deleteSelection();
    const splitPosition = preflight.selection.from;
    if (!canSplit(preflight.doc, splitPosition)) {
      return false;
    }
    const { rightPosition } = state.selection.empty
      ? splitParagraphWithProperties({
          transaction,
          pos: splitPosition,
          transition: { type: "split-left-created-right-retains" },
        })
      : replaceSelectionThenSplitParagraphWithProperties({
          transaction,
          transition: { type: "split-left-created-right-retains" },
        });
    transaction.setSelection(TextSelection.near(transaction.doc.resolve(rightPosition + 1), 1));
    splitResult.tr = transaction;
    splitResult.ownershipRebound = true;
  } else if (!splitBlock(state, capturingDispatch, view)) {
    return false;
  }

  if (dispatch && splitResult.tr !== null) {
    // After split, cursor is in the new (second) paragraph.
    // Apply attr inheritance, border clearing, and stored marks to the SAME transaction.
    const tr = splitResult.tr;
    const { $from } = tr.selection;
    const newPara = $from.parent;

    if (newPara.type.name === "paragraph") {
      let projectedNewParagraph = newPara;
      if (
        !splitResult.ownershipRebound &&
        sourcePara &&
        sourcePropertyState &&
        sourcePosition !== null
      ) {
        const rightPosition = $from.before();
        const leftParagraph = tr.doc.resolve(rightPosition).nodeBefore;
        if (leftParagraph?.type.name !== "paragraph") {
          panic("A block split lost its left paragraph");
        }
        rebindSplitParagraphProperties({
          transaction: tr,
          leftPosition: rightPosition - leftParagraph.nodeSize,
          rightPosition,
          sourceState: sourcePropertyState,
          transition: { type: "split-left-created-right-retains" },
        });
        const rebound = tr.doc.nodeAt(rightPosition);
        if (rebound?.type.name !== "paragraph") {
          panic("A block split lost its right paragraph");
        }
        projectedNewParagraph = rebound;
      }
      // Word's `w:next`: pressing Enter at the end of a paragraph (the new
      // paragraph is empty) switches it to the style's follow-on style — e.g.
      // a heading drops to body text. Only applies to an empty trailing
      // paragraph; splitting mid-paragraph keeps the style on both halves.
      // Use `content.size === 0` rather than `textContent.length` so a
      // mid-paragraph split before an inline atom (image, equation, field,
      // sdt, shape) is not mistaken for an empty trailing paragraph.
      if (resolver !== null && sourcePara !== null && newPara.content.size === 0) {
        const nextStyleId = resolver.getNextStyleId(
          sourcePara.attrs["styleId"] as string | null | undefined,
        );
        projectSplitParagraph({
          context: sourceTableContext,
          resolver,
          sourcePara,
          state,
          tr,
          newPara: projectedNewParagraph,
          useNextStyle: nextStyleId !== null,
        });
        dispatch(tr.scrollIntoView());
        return true;
      }

      const projectedAttrs = expectParagraphAttrs(projectedNewParagraph);
      const propertyState = expectParagraphPropertyState(projectedAttrs._paragraphPropertyState);
      if (propertyState.authoredPPr.borders !== undefined) {
        applyParagraphPropertyProjection({
          transaction: tr,
          pos: $from.before(),
          projection: transitionParagraphProperties({
            attrs: projectedAttrs,
            state: {
              type: "update",
              authored: {
                type: "mutate",
                mutations: [{ key: "borders", mutation: { type: "remove" } }],
              },
              context: { type: "preserve" },
            },
          }),
          source: { type: "preserve" },
        });
      }
    }

    dispatch(tr.scrollIntoView());
  }

  return true;
};

export const BaseKeymapExtension = createExtension({
  name: "baseKeymap",
  priority: Priority.Low,
  onSchemaReady(_ctx: ExtensionContext): ExtensionRuntime {
    return {
      keyboardShortcuts: {
        // Base keymap provides default editing commands
        ...baseKeymap,
        // Override some keys with better defaults
        Enter: splitBlockClearBorders,
        Backspace: chainCommands(
          deleteSelectionWithParagraphOwnership,
          clearIndentOnBackspace,
          withParagraphJoinOwnership(joinBackward, "backward"),
        ),
        Delete: chainCommands(
          deleteSelectionWithParagraphOwnership,
          withParagraphJoinOwnership(joinForward, "forward"),
        ),
        "Mod-a": selectAll,
        Escape: selectParentNode,
      },
    };
  },
});
