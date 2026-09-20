/**
 * Base Keymap Extension — wraps prosemirror-commands baseKeymap
 *
 * Priority: Low (150) — must be the last keymap so other extensions can override keys
 */

import {
  baseKeymap,
  splitBlock,
  deleteSelection,
  joinBackward,
  joinForward,
  selectAll,
  selectParentNode,
} from "prosemirror-commands";
import type { Mark, Node as PMNode, Schema } from "prosemirror-model";
import type { Command, Transaction } from "prosemirror-state";

import type { TextFormatting } from "../../../types/document";
import { mergeTextFormatting } from "../../../utils/textFormattingMerge";
import { expectCharacterStyleMarkAttrs, expectRunFormattingOverrideMarkAttrs } from "../../attrs";
import { keepSectionBreaksOnSurvivingMarks } from "../../commands/sectionBreak";
import { getDocumentStyleResolver } from "../../plugins/documentStyles";
import { RUN_FORMATTING_MARK_NAMES } from "../../runFormattingMarkNames";
import { authoredRunFormattingFromAttrs } from "../../runFormattingProvenance";
import { paragraphAttrsFromResolvedStyle } from "../../styles/resolvedStyleAttrs";
import type { StyleResolver } from "../../styles/styleResolver";
import { createExtension } from "../create";
import { marksToTextFormatting, textFormattingToMarks } from "../marks/markUtils";
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

  if (dispatch) {
    const pos = $cursor.before();
    const tr = state.tr.setNodeMarkup(pos, undefined, {
      ...attrs,
      indentFirstLine: null,
      hangingIndent: null,
      indentLeft: null,
    });
    dispatch(tr.scrollIntoView());
  }
  return true;
};

/**
 * Custom Enter handler: splits the block, inherits style-related attrs,
 * clears paragraph borders, and preserves run formatting on the new paragraph.
 *
 * splitBlock creates a new paragraph with default attrs (all null),
 * so we must manually copy style-related attrs from the source paragraph.
 * Word does NOT propagate paragraph borders (w:pBdr) on Enter.
 */
const INHERITED_PARA_ATTRS = [
  "defaultTextFormatting",
  "styleId",
  "_tableOfContentsLevel",
  "lineSpacing",
  "lineSpacingRule",
  "snapToGrid",
  "spaceAfter",
  "spaceBefore",
  "contextualSpacing",
] as const;

/** Style formatting needed when the caret has no marks of its own. */
const STYLE_MARK_NAMES = new Set(["fontFamily", "fontSize", "textColor"]);

const directFormattingFromMarks = (marks: readonly Mark[]): TextFormatting | undefined => {
  const override = marks.find(({ type }) => type.name === "runFormattingOverride");
  const directFormatting = override
    ? authoredRunFormattingFromAttrs(expectRunFormattingOverrideMarkAttrs(override))
    : undefined;
  const characterStyle = marks.find(({ type }) => type.name === "characterStyle");
  if (!characterStyle) {
    return directFormatting;
  }
  return {
    ...directFormatting,
    styleId: expectCharacterStyleMarkAttrs(characterStyle).styleId,
  };
};

const marksForParagraphFormatting = (
  formatting: TextFormatting,
  directFormatting: TextFormatting,
  schema: Schema,
): Mark[] => {
  const marks = textFormattingToMarks(formatting, schema, {
    authoredCarrier: "preserve",
    directFormatting,
    overrideFormatting: directFormatting,
  });
  if (formatting.styleId) {
    marks.push(schema.mark("characterStyle", { styleId: formatting.styleId }));
  }
  return marks;
};

/**
 * If `sourcePara`'s style defines a `w:next`, replace the empty `newPara`
 * with that style's resolved attrs and seed stored marks from its run
 * formatting. Returns true when a switch happened (caller should dispatch
 * the transaction as-is), false when the source style has no `w:next` and
 * the caller should fall back to the regular inheritance path.
 */
function applyNextParagraphStyle(
  tr: Transaction,
  sourcePara: PMNode,
  newPara: PMNode,
  resolver: StyleResolver,
): boolean {
  const nextStyleId = resolver.getNextStyleId(
    sourcePara.attrs["styleId"] as string | null | undefined,
  );
  if (!nextStyleId) {
    return false;
  }

  const resolved = resolver.resolveParagraphStyle(nextStyleId);
  const styleName = resolver.getStyle(nextStyleId)?.name;
  const { $from } = tr.selection;
  // `paragraphAttrsFromResolvedStyle` already projects the next style's
  // borders (or null), which both clears the source paragraph's leftover
  // border and applies a bordered next style (callouts, etc.).
  tr.setNodeMarkup($from.before(), undefined, {
    ...newPara.attrs,
    styleId: nextStyleId,
    ...paragraphAttrsFromResolvedStyle(resolved, {
      styleId: nextStyleId,
      ...(styleName ? { styleName } : {}),
    }),
  });

  // setStoredMarks MUST come after setNodeMarkup — every step clears it.
  tr.setStoredMarks(
    resolved.runFormatting ? textFormattingToMarks(resolved.runFormatting, tr.doc.type.schema) : [],
  );
  return true;
}

export const splitBlockClearBorders: Command = (state, dispatch, view) => {
  // Capture source paragraph info BEFORE split (splitBlock resets everything)
  const { $from: preSplitFrom } = state.selection;
  const sourcePara = preSplitFrom.parent.type.name === "paragraph" ? preSplitFrom.parent : null;

  // Collect run formatting from the cursor position before splitting.
  // Use storedMarks if set, otherwise resolve from the position.
  const preMarks = state.storedMarks || preSplitFrom.marks();
  const caretFormattingMarks = preMarks.filter((mark) =>
    RUN_FORMATTING_MARK_NAMES.has(mark.type.name),
  );
  const resolver = getDocumentStyleResolver(state);

  // Intercept splitBlock's transaction so we can modify it before dispatch.
  // This ensures attrs + stored marks are set in a single transaction,
  // avoiding a flash where the empty paragraph has no formatting.
  const splitResult = { tr: null as Transaction | null };
  const capturingDispatch = dispatch
    ? (tr: Transaction) => {
        splitResult.tr = tr;
      }
    : undefined;

  if (!splitBlock(state, capturingDispatch, view)) {
    return false;
  }

  if (dispatch && splitResult.tr !== null) {
    // After split, cursor is in the new (second) paragraph.
    // Apply attr inheritance, border clearing, and stored marks to the SAME transaction.
    const tr = splitResult.tr;
    const { $from } = tr.selection;
    const newPara = $from.parent;

    if (newPara.type.name === "paragraph") {
      // Word's `w:next`: pressing Enter at the end of a paragraph (the new
      // paragraph is empty) switches it to the style's follow-on style — e.g.
      // a heading drops to body text. Only applies to an empty trailing
      // paragraph; splitting mid-paragraph keeps the style on both halves.
      // Use `content.size === 0` rather than `textContent.length` so a
      // mid-paragraph split before an inline atom (image, equation, field,
      // sdt, shape) is not mistaken for an empty trailing paragraph.
      if (
        resolver !== null &&
        sourcePara !== null &&
        newPara.content.size === 0 &&
        applyNextParagraphStyle(tr, sourcePara, newPara, resolver)
      ) {
        const directFormatting = directFormattingFromMarks(caretFormattingMarks);
        if (directFormatting && Object.keys(directFormatting).length > 0) {
          const nextParagraph = tr.selection.$from.parent;
          const nextStyleFormatting = nextParagraph.attrs["defaultTextFormatting"] as
            | TextFormatting
            | undefined;
          const defaultTextFormatting =
            mergeTextFormatting(nextStyleFormatting, directFormatting) ?? directFormatting;
          tr.setNodeMarkup(tr.selection.$from.before(), undefined, {
            ...nextParagraph.attrs,
            defaultTextFormatting,
          });
          tr.setStoredMarks(
            marksForParagraphFormatting(defaultTextFormatting, directFormatting, state.schema),
          );
        }
        dispatch(tr.scrollIntoView());
        return true;
      }

      const newAttrs = { ...newPara.attrs };
      let attrsChanged = false;

      // Copy inherited attrs from source paragraph
      if (sourcePara) {
        for (const key of INHERITED_PARA_ATTRS) {
          const srcVal = sourcePara.attrs[key];
          if (srcVal !== null && newAttrs[key] === null) {
            newAttrs[key] = srcVal;
            attrsChanged = true;
          }
        }
      }

      // Clear borders (Word does not propagate paragraph borders on Enter)
      if (newAttrs["borders"]) {
        newAttrs["borders"] = null;
        attrsChanged = true;
      }

      if (attrsChanged) {
        tr.setNodeMarkup($from.before(), undefined, newAttrs);
      }

      // For empty paragraphs (Enter at end of line), preserve direct run formatting
      // so the next typed text keeps the caret's formatting.
      if (newPara.textContent.length === 0) {
        // Determine effective style marks. When text has explicit marks (e.g. user
        // applied a font override), use those. When text inherits formatting from
        // the paragraph style chain (no explicit marks), derive marks from the
        // source paragraph's defaultTextFormatting.
        let effectiveMarks: Mark[] = caretFormattingMarks;

        if (effectiveMarks.length === 0 && sourcePara) {
          const dtf = sourcePara.attrs["defaultTextFormatting"] as TextFormatting | undefined;
          if (dtf) {
            const allMarks = textFormattingToMarks(dtf, state.schema);
            effectiveMarks = allMarks.filter((m) => STYLE_MARK_NAMES.has(m.type.name));
          }
        }

        if (effectiveMarks.length > 0) {
          // Persist every run-formatting mark on the paragraph itself. Stored
          // marks disappear when the selection moves, so font-only defaults
          // would lose direct bold, italic, color, and similar choices after
          // leaving and returning to this empty paragraph.
          const previousFormatting = newAttrs["defaultTextFormatting"] as
            | TextFormatting
            | undefined;
          const markFormatting = marksToTextFormatting(effectiveMarks);
          const defaultTextFormatting =
            mergeTextFormatting(previousFormatting, markFormatting) ?? markFormatting;
          tr.setNodeMarkup($from.before(), undefined, {
            ...newAttrs,
            defaultTextFormatting,
          });

          // IMPORTANT: setStoredMarks MUST be called AFTER all setNodeMarkup calls.
          // setNodeMarkup adds a ReplaceStep which clears storedMarks on the transaction.
          tr.setStoredMarks(effectiveMarks);
        }
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
        // A join consumes one paragraph mark and keeps the other; the section
        // break follows the mark that survived rather than the node whose
        // attrs ProseMirror happened to keep.
        Backspace: keepSectionBreaksOnSurvivingMarks(
          chainCommands(deleteSelection, clearIndentOnBackspace, joinBackward),
        ),
        Delete: keepSectionBreaksOnSurvivingMarks(chainCommands(deleteSelection, joinForward)),
        "Mod-a": selectAll,
        Escape: selectParentNode,
      },
    };
  },
});
