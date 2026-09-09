/** Document-scoped numbering definitions for list-editing commands. */

import { Plugin, PluginKey, type EditorState } from "prosemirror-state";

import { getCachedNumberingMap, type NumberingMap } from "../../docx/numberingParser";
import type { NumberingDefinitions } from "../../types/document";

export const documentNumberingKey = new PluginKey<NumberingMap | null>("documentNumbering");

export const createDocumentNumberingPlugin = (
  definitions: NumberingDefinitions | null | undefined,
): Plugin<NumberingMap | null> => {
  const numbering = definitions ? getCachedNumberingMap(definitions) : null;
  return new Plugin({
    key: documentNumberingKey,
    state: {
      init: () => numbering,
      apply: (_transaction, value) => value,
    },
  });
};

export const getDocumentNumbering = (state: EditorState): NumberingMap | null =>
  documentNumberingKey.getState(state) ?? null;
