/**
 * The numbering references a ProseMirror document makes, read with the
 * rendering each one resolved to. {@link completeListNumbering} defines from
 * these the instances a list command minted and the package does not have yet.
 */

import type { Node as PMNode } from "prosemirror-model";

import { completeListNumbering, type ListInstanceReference } from "../docx/listNumberingInstances";
import { effectiveParagraphNumberingReference } from "./numberingAttr";
import type { NumberingDefinitions } from "../types/document";
import { expectParagraphAttrs } from "./attrs";
import {
  hasListRendering,
  listRenderingFromAttrs,
  type ListRenderingSourceAttrs,
} from "./listRenderingAttrs";

const PARAGRAPH_NODE = "paragraph";

type ReferenceSink = {
  isDefined: (numId: number) => boolean;
  references: ListInstanceReference[];
};

/**
 * Record `attrs`' reference when it names an instance the package lacks and
 * states what it renders as. A reference with no rendering carries nothing to
 * define an instance from.
 */
const collectReference = (attrs: ListRenderingSourceAttrs, sink: ReferenceSink): void => {
  const numbering = effectiveParagraphNumberingReference(attrs);
  if (numbering === undefined || sink.isDefined(numbering.numId)) {
    return;
  }
  if (!hasListRendering(attrs)) {
    return;
  }
  sink.references.push({
    numId: numbering.numId,
    ilvl: numbering.ilvl,
    rendering: listRenderingFromAttrs({ attrs, numId: numbering.numId }),
  });
};

/**
 * Every reference in `doc`, live or recorded as a tracked change's previous
 * state, to an instance `isDefined` does not know. A rejected change restores
 * its previous numbering, so that instance has to exist as well.
 */
const undefinedListInstanceReferences = (
  doc: PMNode,
  isDefined: (numId: number) => boolean,
): ListInstanceReference[] => {
  const sink: ReferenceSink = { isDefined, references: [] };
  doc.descendants((node) => {
    if (node.type.name !== PARAGRAPH_NODE) {
      return true;
    }
    const attrs = expectParagraphAttrs(node);
    collectReference(attrs, sink);
    for (const change of attrs._propertyChanges ?? []) {
      const previous = change.previousFormatting;
      if (previous) {
        collectReference(previous, sink);
      }
    }
    return false;
  });
  return sink.references;
};

/** The instances one paragraph names, live or as a tracked change's previous state. */
export const paragraphListReferences = (node: PMNode): number[] => {
  if (node.type.name !== PARAGRAPH_NODE) {
    return [];
  }
  const attrs = expectParagraphAttrs(node);
  const numIds: number[] = [];
  const add = (source: ListRenderingSourceAttrs): void => {
    const numbering = effectiveParagraphNumberingReference(source);
    if (numbering !== undefined) numIds.push(numbering.numId);
  };
  add(attrs);
  for (const change of attrs._propertyChanges ?? []) {
    const previous = change.previousFormatting;
    if (previous) add(previous);
  }
  return numIds;
};

/**
 * `definitions` plus every instance `doc` references and only its paragraphs
 * define: what the package's numbering part has to hold for `doc` to save.
 */
export const completeNumberingForDoc = (
  definitions: NumberingDefinitions | undefined,
  doc: PMNode,
): NumberingDefinitions | undefined => {
  const defined = new Set((definitions?.nums ?? []).map(({ numId }) => numId));
  return completeListNumbering(
    definitions,
    undefinedListInstanceReferences(doc, (numId) => defined.has(numId)),
  );
};
