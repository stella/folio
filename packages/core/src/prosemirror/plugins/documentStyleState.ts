import { PluginKey, type EditorState } from "prosemirror-state";
import { type BuiltInStyleIndex, EMPTY_BUILT_IN_STYLE_INDEX } from "../../docx/builtInStyles";
import type { Style } from "../../types/document";
import type { StyleResolver } from "../styles/styleResolver";

export const documentStylesKey = new PluginKey<StyleResolver | null>("documentStyles");

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
