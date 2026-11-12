/** Document-scoped numbering definitions for list-editing commands. */

import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, Plugin, type Transaction } from "prosemirror-state";
import {
  AddMarkStep,
  DocAttrStep,
  RemoveMarkStep,
  ReplaceAroundStep,
  ReplaceStep,
  type Step,
} from "prosemirror-transform";

import {
  getCachedNumberingMap,
  type NumberingMap,
} from "../../docx/numberingParser";
import type { NumberingDefinitions } from "../../types/document";
import {
  completeNumberingForDoc,
  paragraphListReferences,
} from "../listInstanceReferences";
import { createPluginKeyOwner } from "../pluginKeyOwner";

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

const documentNumberingKey =
  createPluginKeyOwner<DocumentNumberingState>("documentNumbering");

const completedState = (
  definitions: NumberingDefinitions | null,
  doc: PMNode,
): DocumentNumberingState => {
  const completed = completeNumberingForDoc(definitions ?? undefined, doc);
  return {
    definitions,
    map: completed ? getCachedNumberingMap(completed) : null,
  };
};

const isReplace = (step: Step): step is ReplaceStep | ReplaceAroundStep =>
  step instanceof ReplaceStep || step instanceof ReplaceAroundStep;

const leavesNumberingUnchanged = (step: Step): boolean =>
  step instanceof AddMarkStep ||
  step instanceof RemoveMarkStep ||
  step instanceof DocAttrStep;

/** The instances a package's numbering part defines, by its definitions object. */
const definedInstances = new WeakMap<
  NumberingDefinitions,
  ReadonlySet<number>
>();
const instancesDefinedBy = (
  definitions: NumberingDefinitions | null,
): ReadonlySet<number> => {
  if (definitions === null) return new Set();
  let defined = definedInstances.get(definitions);
  if (defined === undefined) {
    defined = new Set((definitions.nums ?? []).map(({ numId }) => numId));
    definedInstances.set(definitions, defined);
  }
  return defined;
};

/**
 * Whether a paragraph names an instance the package does not define: only
 * those reach the completed numbering, which adds an instance for each.
 */
const namesUndefinedInstance = (
  node: PMNode,
  defined: ReadonlySet<number>,
): boolean =>
  paragraphListReferences(node).some((numId) => !defined.has(numId));

const containsUndefinedReference = (
  node: PMNode,
  from: number,
  to: number,
  defined: ReadonlySet<number>,
): boolean => {
  let found = false;
  node.nodesBetween(from, to, (child, pos) => {
    if (
      child.type.name !== "paragraph" ||
      pos < from ||
      pos + child.nodeSize > to ||
      !namesUndefinedInstance(child, defined)
    ) {
      return !found;
    }
    found = true;
    return false;
  });
  return found;
};

/**
 * Whether a transaction can change the completed numbering: the instances the
 * document names that the package does not define. Text and mark edits do
 * not; paragraph attrs can. A structural replacement can only when a
 * paragraph it inserts, removes or joins names such an instance, so typing in
 * a list whose instances the package defines leaves the numbering as it is.
 * Replacements are inspected in the step-local document rather than by
 * walking the completed document.
 */
const changesNumberingInputs = (
  tr: Transaction,
  defined: ReadonlySet<number>,
): boolean =>
  tr.steps.some((step, index) => {
    if (leavesNumberingUnchanged(step)) return false;
    if (!isReplace(step)) return true;

    let insertsUndefinedReference = false;
    step.slice.content.descendants((node) => {
      if (!namesUndefinedInstance(node, defined))
        return !insertsUndefinedReference;
      insertsUndefinedReference = true;
      return false;
    });
    if (insertsUndefinedReference) return true;

    const before = tr.docs[index];
    if (before === undefined) return true;
    const fromParent = before.resolve(step.from).parent;
    const toParent = before.resolve(step.to).parent;
    if (
      fromParent !== toParent &&
      (namesUndefinedInstance(fromParent, defined) ||
        namesUndefinedInstance(toParent, defined))
    ) {
      return true;
    }
    return containsUndefinedReference(before, step.from, step.to, defined);
  });

export const createDocumentNumberingPlugin = (
  definitions: NumberingDefinitions | null | undefined,
): Plugin =>
  new Plugin<DocumentNumberingState>({
    key: documentNumberingKey.createKey(),
    state: {
      init: (_config, state) => completedState(definitions ?? null, state.doc),
      apply: (tr, current, _oldState, state) =>
        tr.docChanged &&
        changesNumberingInputs(tr, instancesDefinedBy(current.definitions))
          ? completedState(current.definitions, state.doc)
          : current,
    },
  });

/** Numbering for states without the plugin, one per document node. */
const untrackedNumbering = new WeakMap<PMNode, DocumentNumberingState>();

/** Numbering derived from a document when no editor plugin owns package definitions. */
export const getDocumentNumberingForDoc = (
  doc: PMNode,
): NumberingMap | null => {
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
export const getDocumentNumbering = (
  state: EditorState,
): NumberingMap | null => {
  const tracked = documentNumberingKey.getState(state);
  if (tracked) {
    return tracked.map;
  }
  // A plugin applying a transaction sees only the states of the plugins
  // before it. Read there, numbering placed later would look absent and be
  // rebuilt from the paragraphs alone, without the package's definitions.
  if (documentNumberingKey.get(state) !== undefined)
    panic(
      "Document numbering was read before its plugin applied: place the numbering plugin before the plugins that read it.",
    );
  // A state without the plugin knows no package numbering; its paragraphs
  // still define the lists commands started in it.
  return getDocumentNumberingForDoc(state.doc);
};

/** The package's own definitions, without the instances the document defines. */
export const getPackageNumberingDefinitions = (
  state: EditorState,
): NumberingDefinitions | null =>
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
export const getDocumentNumberingInstanceIds = (
  state: EditorState,
): ReadonlySet<number> | null => {
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
    plugins: [createDocumentNumberingPlugin(definitions), ...retained],
  });
};
