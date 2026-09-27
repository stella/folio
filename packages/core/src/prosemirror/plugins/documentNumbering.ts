/** Document-scoped numbering definitions for list-editing commands. */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey, type EditorState, type Transaction } from "prosemirror-state";
import {
  AddMarkStep,
  DocAttrStep,
  RemoveMarkStep,
  ReplaceAroundStep,
  ReplaceStep,
  type Step,
} from "prosemirror-transform";

import { getCachedNumberingMap, type NumberingMap } from "../../docx/numberingParser";
import type { NumberingDefinitions } from "../../types/document";
import { completeNumberingForDoc, paragraphListReferences } from "../listInstanceReferences";

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

const documentNumberingKey = new PluginKey<DocumentNumberingState>("documentNumbering");

const completedState = (
  definitions: NumberingDefinitions | null,
  doc: PMNode,
): DocumentNumberingState => {
  const completed = completeNumberingForDoc(definitions ?? undefined, doc);
  return { definitions, map: completed ? getCachedNumberingMap(completed) : null };
};

/** Steps that can change a paragraph's attrs or bring paragraphs in, by what they replace. */
const isReplace = (step: Step): step is ReplaceStep | ReplaceAroundStep =>
  step instanceof ReplaceStep || step instanceof ReplaceAroundStep;

/** Steps that change nothing a numbering reference lives in. */
const leavesParagraphAttrs = (step: Step): boolean =>
  step instanceof AddMarkStep || step instanceof RemoveMarkStep || step instanceof DocAttrStep;

/**
 * Whether `tr` may have introduced a reference `map` does not define. Only the
 * ranges its steps replaced are read, so typing costs one paragraph, not the
 * document. A step that changes attrs in place (not by replacement) is not
 * localised and answers yes.
 */
const mayReferenceUnknownInstance = (tr: Transaction, map: NumberingMap | null): boolean => {
  const ranges: { from: number; to: number }[] = [];
  for (const [index, step] of tr.steps.entries()) {
    if (leavesParagraphAttrs(step)) {
      continue;
    }
    if (!isReplace(step)) {
      return true;
    }
    const later = tr.mapping.slice(index + 1);
    step.getMap().forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      ranges.push({ from: later.map(newStart, -1), to: later.map(newEnd, 1) });
    });
  }
  const size = tr.doc.content.size;
  const isUnknown = (numId: number): boolean => !(map?.hasNumbering(numId) ?? false);
  return ranges.some(({ from, to }) => {
    let unknown = false;
    const start = Math.min(Math.max(0, from), size);
    tr.doc.nodesBetween(start, Math.min(Math.max(start, to), size), (node) => {
      if (unknown) {
        return false;
      }
      if (node.type.name !== "paragraph") {
        return true;
      }
      unknown = paragraphListReferences(node).some(isUnknown);
      return false;
    });
    return unknown;
  });
};

export const createDocumentNumberingPlugin = (
  definitions: NumberingDefinitions | null | undefined,
): Plugin =>
  new Plugin<DocumentNumberingState>({
    key: documentNumberingKey,
    state: {
      init: (_config, state) => completedState(definitions ?? null, state.doc),
      apply: (tr, current, _oldState, state) =>
        tr.docChanged && mayReferenceUnknownInstance(tr, current.map)
          ? completedState(current.definitions, state.doc)
          : current,
    },
  });

/** Numbering for states without the plugin, one per document node. */
const untrackedNumbering = new WeakMap<PMNode, DocumentNumberingState>();

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
  let untracked = untrackedNumbering.get(state.doc);
  if (!untracked) {
    untracked = completedState(null, state.doc);
    untrackedNumbering.set(state.doc, untracked);
  }
  return untracked.map;
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
