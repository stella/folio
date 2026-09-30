/**
 * documentStyles plugin — makes the document's StyleResolver reachable from
 * ProseMirror commands.
 *
 * Styles otherwise flow one way (Document → PM) at load time: `toProseDoc`
 * bakes resolved formatting into nodes and discards the resolver. Some
 * commands need the live style table though — the Enter handler looks up a
 * paragraph style's `w:next` to switch to body text after a heading. This
 * plugin parks the resolver in plugin state so those commands can read it
 * via `getDocumentStyleResolver(state)`.
 *
 * The host (HiddenProseMirror / HiddenHeaderFooterPMs) passes the same
 * styles it hands to `toProseDoc` and adds this plugin when creating the
 * EditorState. When absent, style-aware commands fall back to their
 * style-agnostic behavior.
 */

import { Mark } from "prosemirror-model";
import { isHistoryTransaction } from "prosemirror-history";
import { Mapping, ReplaceStep } from "prosemirror-transform";
import { rebaseParagraphRuns } from "../rebaseParagraphRuns";
import { JOINED_RUNS_RESTYLED_META } from "../extensions/features/JoinedRunStyleExtension";
import { Plugin, type EditorState, type Transaction } from "prosemirror-state";

import type { StyleDefinitions } from "../../types/document";
import { StyleResolver, createStyleResolver } from "../styles/styleResolver";
import { documentStylesKey } from "./documentStyleState";

export {
  documentStylesKey,
  getDocumentStyleResolver,
  getDocumentStyleDefinitions,
  getDocumentBuiltInStyles,
} from "./documentStyleState";
import { getDocumentNumbering } from "./documentNumbering";
import { resolveEditedParagraphStyles } from "./paragraphStyleResolution";

/** Rebase surviving right-hand runs when ordinary replacement steps consume a paragraph opening. */
const rebaseEditedParagraphJoins = (
  transactions: readonly Transaction[],
  tr: Transaction,
  resolver: StyleResolver | null,
): void => {
  const maps = transactions.flatMap((transaction) => transaction.mapping.maps);
  let offset = 0;
  for (const transaction of transactions) {
    // Revision resolution and formatting commands already own their run rebases.
    if (
      isHistoryTransaction(transaction) ||
      transaction.getMeta(JOINED_RUNS_RESTYLED_META) === true ||
      !transaction.steps.every((step) => step instanceof ReplaceStep)
    ) {
      offset += transaction.steps.length;
      continue;
    }
    for (const [index, step] of transaction.steps.entries()) {
      const before = transaction.docs[index];
      if (!before) continue;
      const map = step.getMap();
      const following = new Mapping(maps.slice(offset + index + 1));
      const mapping = new Mapping([map, ...following.maps]);
      // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror StepMap.forEach
      map.forEach((oldFrom, oldTo) => {
        before.nodesBetween(oldFrom, oldTo, (paragraph, position) => {
          if (paragraph.type.name !== "paragraph") return true;
          if (!map.mapResult(position, 1).deleted) return false;
          const from = mapping.map(position + 1, 1);
          const to = mapping.map(position + paragraph.nodeSize - 1, -1);
          if (from >= to) return false;
          const at = tr.doc.resolve(from);
          if (at.parent.type.name !== "paragraph") return false;
          const targetPosition = at.start() - 1;
          if (targetPosition === mapping.map(position, 1)) return false;
          const surviving = tr.doc.nodeAt(from);
          const original = before.nodeAt(mapping.invert().map(from, 1));
          if (!surviving || !original || !Mark.sameSet(surviving.marks, original.marks))
            return false;
          rebaseParagraphRuns({
            tr,
            position: targetPosition,
            previous: paragraph,
            styleResolver: resolver,
            range: { from, to },
          });
          return false;
        });
      });
    }
    offset += transaction.steps.length;
  }
};

let bareStyleResolver: StyleResolver | undefined;
/** The cascade of a package that defines no styles: docDefaults-free built-ins. */
const bareResolver = (): StyleResolver => {
  bareStyleResolver ??= createStyleResolver(undefined);
  return bareStyleResolver;
};

/**
 * Create the plugin holding a StyleResolver for the document's `styles` for
 * the lifetime of the EditorState. The resolver is fixed per document load;
 * loading a new document recreates the state (and thus this plugin) with a
 * fresh resolver. Accepts a pre-built resolver too, for callers that already
 * have one.
 */
export function createDocumentStylesPlugin(
  styles: StyleDefinitions | StyleResolver | null | undefined,
): Plugin {
  let resolver: StyleResolver | null;
  if (styles instanceof StyleResolver) {
    resolver = styles;
  } else if (styles) {
    resolver = createStyleResolver(styles);
  } else {
    resolver = null;
  }
  return new Plugin<StyleResolver | null>({
    key: documentStylesKey,
    state: {
      init: () => resolver,
      apply: (_tr, value) => value,
    },
    // A paragraph an edit creates reads its style cascade the way a loaded
    // one does, so it paints its inherited formatting before any reopen.
    // A package without a styles part loads against the built-in defaults.
    appendTransaction: (transactions, _oldState, newState) => {
      const styleResolver = documentStylesKey.getState(newState) ?? bareResolver();
      const tr =
        resolveEditedParagraphStyles(transactions, newState, styleResolver, () =>
          getDocumentNumbering(newState),
        ) ?? newState.tr;
      rebaseEditedParagraphJoins(transactions, tr, styleResolver);
      return tr.docChanged ? tr : null;
    },
  });
}

/** Reconfigure so ProseMirror replaces the keyed plugin's resolver state. */
export const withDocumentStyles = (
  state: EditorState,
  styles: StyleDefinitions | StyleResolver | null | undefined,
): EditorState => {
  const previous = documentStylesKey.get(state);
  const retained = state.plugins.filter((plugin) => plugin !== previous);
  return state.reconfigure({ plugins: retained }).reconfigure({
    plugins: [...retained, createDocumentStylesPlugin(styles)],
  });
};
