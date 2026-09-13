import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { Document, SectionProperties } from "../types/document";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import type { RemovedSectionReference } from "./sectionEndpointResolution";

export type SectionReferenceInventory = {
  references: readonly RemovedSectionReference[];
  revisionRelationships: ReadonlySet<string>;
};

const relationshipKey = ({part, relationshipId}: RemovedSectionReference): string =>
  `${part}:${relationshipId}`;
const referenceKey = (reference: RemovedSectionReference): string =>
  `${relationshipKey(reference)}:${reference.type}`;

export const captureSectionReferenceInventory = (
  document: PMNode,
  finalProperties: SectionProperties | undefined,
): SectionReferenceInventory => {
  const references: RemovedSectionReference[] = [];
  const revisionRelationships = new Set<string>();
  const collect = (properties: SectionProperties | undefined): void => {
    if (!properties) return;
    const start = references.length;
    const append = (selection: Pick<SectionProperties, "headerReferences" | "footerReferences">): void => {
      for (const {type, rId} of selection.headerReferences ?? []) references.push({part: "header", type, relationshipId: rId});
      for (const {type, rId} of selection.footerReferences ?? []) references.push({part: "footer", type, relationshipId: rId});
    };
    append(properties);
    const history = properties.propertyChanges?.filter(({previousReferences}) => previousReferences !== undefined) ?? [];
    for (const change of history) if (change.previousReferences) append(change.previousReferences);
    if (history.length > 0) {
      for (const reference of references.slice(start)) revisionRelationships.add(relationshipKey(reference));
    }
  };
  document.descendants((node) => {
    if (node.type.name === "paragraph") collect(expectParagraphAttrs(node)._sectionProperties);
  });
  collect(finalProperties);
  return {references, revisionRelationships};
};

export const resolvedSectionReferenceLosses = ({
  before,
  after,
}: {before: SectionReferenceInventory; after: SectionReferenceInventory}): RemovedSectionReference[] => {
  const remaining = new Map<string, number>();
  for (const reference of after.references) {
    const key = referenceKey(reference);
    remaining.set(key, (remaining.get(key) ?? 0) + 1);
  }
  const removed: RemovedSectionReference[] = [];
  for (const reference of before.references) {
    if (!before.revisionRelationships.has(relationshipKey(reference))) continue;
    const key = referenceKey(reference);
    const count = remaining.get(key) ?? 0;
    if (count > 0) remaining.set(key, count - 1);
    else removed.push(reference);
  }
  return removed;
};

const activeResolutions = new WeakMap<Document, readonly RemovedSectionReference[]>();

export const withSectionReferenceResolution = async <T>({
  document,
  removedReferences,
  repack,
}: {document: Document; removedReferences: readonly RemovedSectionReference[]; repack: () => Promise<T>}): Promise<T> => {
  if (activeResolutions.has(document)) panic("A section reference repack is already active");
  activeResolutions.set(document, removedReferences.map((reference) => ({...reference})));
  try {
    const result = await repack();
    if (activeResolutions.has(document)) panic("The section reference repack did not consume its resolution");
    return result;
  } finally {
    activeResolutions.delete(document);
  }
};

export const consumeSectionReferenceResolution = (document: Document): readonly RemovedSectionReference[] => {
  const references = activeResolutions.get(document) ?? [];
  activeResolutions.delete(document);
  return references;
};
