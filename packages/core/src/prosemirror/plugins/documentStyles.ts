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

import { Plugin, PluginKey, type EditorState } from "prosemirror-state";

import { type BuiltInStyleIndex, EMPTY_BUILT_IN_STYLE_INDEX } from "../../docx/builtInStyles";
import type { Style, StyleDefinitions } from "../../types/document";
import { StyleResolver, createStyleResolver } from "../styles/styleResolver";
import { getDocumentNumbering } from "./documentNumbering";
import { resolveEditedParagraphStyles } from "./paragraphStyleResolution";

export const documentStylesKey = new PluginKey<StyleResolver | null>("documentStyles");

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
    appendTransaction: (transactions, _oldState, newState) =>
      resolveEditedParagraphStyles(
        transactions,
        newState,
        documentStylesKey.getState(newState) ?? bareResolver(),
        () => getDocumentNumbering(newState),
      ),
  });
}

/** Read the document's StyleResolver, or null when the plugin isn't installed. */
export function getDocumentStyleResolver(state: EditorState): StyleResolver | null {
  return documentStylesKey.getState(state) ?? null;
}

/**
 * The document's style definitions by id, or `null` when the state carries no
 * styles plugin and so cannot say what the document defines. A document
 * without a styles part defines none.
 */
export function getDocumentStyleDefinitions(
  state: EditorState,
): { get: (styleId: string) => Style | undefined; paragraphStyles: () => Style[] } | null {
  if (documentStylesKey.get(state) === undefined) {
    return null;
  }
  const resolver = getDocumentStyleResolver(state);
  return {
    get: (styleId) => resolver?.getStyle(styleId),
    paragraphStyles: () => resolver?.getParagraphStyles() ?? [],
  };
}

/**
 * Read the document's built-in style index, for the callers that classify
 * paragraphs (heading collection, markdown export, the AI snapshot). Without
 * the plugin every lookup misses, which degrades classification to the
 * paragraph's own outline level rather than to English style ids.
 */
export function getDocumentBuiltInStyles(state: EditorState): BuiltInStyleIndex {
  return getDocumentStyleResolver(state)?.builtInStyles ?? EMPTY_BUILT_IN_STYLE_INDEX;
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
