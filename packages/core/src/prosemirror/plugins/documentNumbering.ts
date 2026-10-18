/** Document-scoped numbering definitions for list-editing commands. */

import { createPluginKeyOwner } from "../pluginKeyOwner";
import type { Node as PMNode } from "prosemirror-model";
import { Plugin, type EditorState } from "prosemirror-state";

import { getCachedNumberingMap, type NumberingMap } from "../../docx/numberingParser";
import type { NumberingDefinitions } from "../../types/document";
import { completeNumberingForDoc } from "../listInstanceReferences";

/**
 * The package's own definitions (`null` when it has no numbering part) and
 * the map the document's paragraphs resolve against: those definitions plus
 * every instance a list command defined on the paragraphs it numbered (see
 * `docx/listNumberingInstances.ts`).
 */
type DocumentNumberingState = {
  definitions: NumberingDefinitions | null;
  map: NumberingMap | null;
};

const documentNumberingKey = createPluginKeyOwner<DocumentNumberingState>("documentNumbering");

const completedState = (
  definitions: NumberingDefinitions | null,
  doc: PMNode,
): DocumentNumberingState => {
  const completed = completeNumberingForDoc(definitions ?? undefined, doc);
  return { definitions, map: completed ? getCachedNumberingMap(completed) : null };
};

export const createDocumentNumberingPlugin = (
  definitions: NumberingDefinitions | null | undefined,
): Plugin =>
  new Plugin<DocumentNumberingState>({
    key: documentNumberingKey.createKey(),
    state: {
      init: (_config, state) => completedState(definitions ?? null, state.doc),
      apply: (tr, current, _oldState, state) =>
        tr.docChanged ? completedState(current.definitions, state.doc) : current,
    },
  });

/** Numbering for states without the plugin, one per document node. */
const untrackedNumbering = new WeakMap<PMNode, DocumentNumberingState>();

/** Numbering derived from a document when no editor plugin owns package definitions. */
export const getDocumentNumberingForDoc = (doc: PMNode): NumberingMap | null => {
  let untracked = untrackedNumbering.get(doc);
  if (!untracked) {
    untracked = completedState(null, doc);
    untrackedNumbering.set(doc, untracked);
  }
  return untracked.map;
};

/**
 * The numbering the document's paragraphs resolve against: the package's
 * definitions plus every instance a list command defined in this document.
 * An instance whose paragraphs were all deleted may linger until the next
 * completion; the save completes from the document itself.
 */
export const getDocumentNumbering = (state: EditorState): NumberingMap | null => {
  const tracked = documentNumberingKey.getState(state);
  if (tracked) {
    return tracked.map;
  }
  // A state without the plugin knows no package numbering; its paragraphs
  // still define the lists commands started in it.
  return getDocumentNumberingForDoc(state.doc);
};

/** The package's own definitions, without the instances the document defines. */
export const getPackageNumberingDefinitions = (state: EditorState): NumberingDefinitions | null =>
  documentNumberingKey.getState(state)?.definitions ?? null;

/**
 * Whether the state knows its document's numbering at all. A state without
 * the plugin cannot tell a package with no numbering part from one it was
 * never told about.
 */
export const hasDocumentNumbering = (state: EditorState): boolean =>
  documentNumberingKey.get(state) !== undefined;

/**
 * The document's numbering definitions: `null` for a document without a
 * numbering part, `undefined` when the state carries no numbering plugin and
 * so cannot say what the document defines.
 */
export const getStatedDocumentNumbering = (
  state: EditorState,
): NumberingDefinitions | null | undefined => {
  if (documentNumberingKey.get(state) === undefined) {
    return undefined;
  }
  return getDocumentNumbering(state)?.definitions ?? null;
};

/**
 * The numbering instances (`w:num` ids) the document defines, or `null` when
 * the state carries no numbering plugin and so cannot say. A document without
 * a numbering part defines none.
 */
export const getDocumentNumberingInstanceIds = (state: EditorState): ReadonlySet<number> | null => {
  const numbering = getStatedDocumentNumbering(state);
  if (numbering === undefined) {
    return null;
  }
  return new Set(numbering?.nums.map(({ numId }) => numId) ?? []);
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
