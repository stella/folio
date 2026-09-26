/** Document-scoped numbering definitions for list-editing commands. */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey, type EditorState } from "prosemirror-state";

import { getCachedNumberingMap, type NumberingMap } from "../../docx/numberingParser";
import type { NumberingDefinitions } from "../../types/document";
import { completeNumberingForDoc } from "../listInstanceReferences";

/** The package's own definitions; `null` when it has no numbering part. */
type DocumentNumberingState = {
  definitions: NumberingDefinitions | null;
};

const documentNumberingKey = new PluginKey<DocumentNumberingState>("documentNumbering");

export const createDocumentNumberingPlugin = (
  definitions: NumberingDefinitions | null | undefined,
): Plugin => {
  const value: DocumentNumberingState = { definitions: definitions ?? null };
  return new Plugin<DocumentNumberingState>({
    key: documentNumberingKey,
    state: {
      init: () => value,
      apply: (_transaction, current) => current,
    },
  });
};

/**
 * One completed map per document node and package definitions. A list command
 * that defines a new instance records it only on the paragraphs it numbers
 * (see `docx/listNumberingInstances.ts`), so the map a command reads has to
 * include what the document itself defines.
 */
const completedNumbering = new WeakMap<
  PMNode,
  { definitions: NumberingDefinitions | null; map: NumberingMap | null }
>();

/**
 * The numbering the document's paragraphs resolve against: the package's
 * definitions plus every instance a list command defined in this document.
 */
export const getDocumentNumbering = (state: EditorState): NumberingMap | null => {
  const definitions = documentNumberingKey.getState(state)?.definitions ?? null;
  const cached = completedNumbering.get(state.doc);
  if (cached && cached.definitions === definitions) {
    return cached.map;
  }
  const completed = completeNumberingForDoc(definitions ?? undefined, state.doc);
  const map = completed ? getCachedNumberingMap(completed) : null;
  completedNumbering.set(state.doc, { definitions, map });
  return map;
};

/** Replace numbering state while retaining every unrelated plugin state. */
export const withDocumentNumbering = (
  state: EditorState,
  definitions: NumberingDefinitions,
): EditorState => {
  const previous = documentNumberingKey.get(state);
  const retained = state.plugins.filter((plugin) => plugin !== previous);
  return state.reconfigure({ plugins: retained }).reconfigure({
    plugins: [...retained, createDocumentNumberingPlugin(definitions)],
  });
};
