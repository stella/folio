import type { Node as PMNode } from "prosemirror-model";

import { sectionReferenceHistory } from "../docx/sectionReferenceHistory";
import type { SectionProperties } from "../types/document";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import { resolveAllChangesInHeadlessStateWithMapping } from "../prosemirror/commands/comments";
import { canonicalJson } from "../utils/canonicalJson";

export type SectionBoundaryPropertyChange = {
  position: number;
  current: SectionProperties;
  target: SectionProperties;
};

export type SectionBoundaryPropertiesComparison =
  | { status: "matched"; changes: readonly SectionBoundaryPropertyChange[] }
  | { status: "unalignable"; detail: string };

type BoundaryParagraph = {
  position: number;
  text: string;
  sectionProperties: SectionProperties | undefined;
};

type BoundaryChild =
  | { type: "paragraph"; paragraph: BoundaryParagraph }
  | { type: "other"; nodeType: string };

const withoutChangeHistory = (properties: SectionProperties): SectionProperties => {
  const { propertyChanges: _propertyChanges, ...result } = properties;
  return result;
};

const boundaryChildrenOf = (document: PMNode): BoundaryChild[] | null => {
  const children: BoundaryChild[] = [];
  let hasChangeHistory = false;
  document.forEach((node, position) => {
    if (node.type.name !== "paragraph") {
      children.push({ type: "other", nodeType: node.type.name });
      return;
    }
    const attrs = expectParagraphAttrs(node);
    const sectionProperties = attrs._sectionProperties;
    if (sectionProperties?.propertyChanges !== undefined) {
      hasChangeHistory = true;
      return;
    }
    children.push({
      type: "paragraph",
      paragraph: { position, text: node.textContent, sectionProperties },
    });
  });
  return hasChangeHistory ? null : children;
};

/**
 * Compare section properties attached to retained top-level paragraph endpoints.
 * A section endpoint cannot be added to or removed from a retained paragraph
 * with `sectPrChange`: its previous payload is still a section endpoint.
 */
export const compareSectionBoundaryProperties = ({
  current,
  target,
  mapTargetProperties,
}: {
  current: PMNode;
  target: PMNode;
  mapTargetProperties?: (properties: SectionProperties) => SectionProperties | null;
}): SectionBoundaryPropertiesComparison => {
  const currentChildren = boundaryChildrenOf(current);
  const targetChildren = boundaryChildrenOf(target);
  if (!currentChildren || !targetChildren) {
    return { status: "unalignable", detail: "section boundary revision history is unsupported" };
  }
  if (currentChildren.length !== targetChildren.length) {
    return { status: "unalignable", detail: "section boundary topology differs" };
  }

  const changes: SectionBoundaryPropertyChange[] = [];
  for (const [index, currentChild] of currentChildren.entries()) {
    const targetChild = targetChildren[index];
    if (!targetChild || currentChild.type !== targetChild.type) {
      return { status: "unalignable", detail: "section boundary topology differs" };
    }
    if (currentChild.type === "other" && targetChild.type === "other") {
      if (currentChild.nodeType !== targetChild.nodeType) {
        return { status: "unalignable", detail: "section boundary topology differs" };
      }
      continue;
    }
    if (currentChild.type !== "paragraph" || targetChild.type !== "paragraph") {
      return { status: "unalignable", detail: "section boundary topology differs" };
    }
    const currentParagraph = currentChild.paragraph;
    const targetParagraph = targetChild.paragraph;
    if (currentParagraph.text !== targetParagraph.text) {
      return { status: "unalignable", detail: "retained paragraph text differs" };
    }
    const currentProperties = currentParagraph.sectionProperties;
    const targetProperties = targetParagraph.sectionProperties;
    if (currentProperties === undefined || targetProperties === undefined) {
      if (currentProperties !== targetProperties) {
        return { status: "unalignable", detail: "section endpoint presence differs" };
      }
      continue;
    }
    const currentWithoutHistory = withoutChangeHistory(currentProperties);
    const mappedTarget = mapTargetProperties
      ? mapTargetProperties(targetProperties)
      : targetProperties;
    if (mappedTarget === null)
      return { status: "unalignable", detail: "section references cannot be mapped" };
    const targetWithoutHistory = withoutChangeHistory(mappedTarget);
    if (canonicalJson(currentWithoutHistory) === canonicalJson(targetWithoutHistory)) continue;
    changes.push({
      position: currentParagraph.position,
      current: currentWithoutHistory,
      target: targetWithoutHistory,
    });
  }
  return { status: "matched", changes };
};

const boundaryParagraphEntriesOf = (document: PMNode): BoundaryParagraph[] | null => {
  const children = boundaryChildrenOf(document);
  if (!children) return null;
  return children.flatMap((child) => (child.type === "paragraph" ? [child.paragraph] : []));
};

type MappedSectionBoundaryProperties =
  | { kind: "inserted"; target: SectionProperties }
  | { kind: "retained"; previous: SectionProperties; target: SectionProperties };

export type StageSectionBoundaryPropertiesResult =
  | {
      status: "matched";
      transaction: import("prosemirror-state").Transaction;
      nextRevisionId: number;
      rangeCount: number;
    }
  | { status: "unalignable"; detail: string }
  | { status: "budget-exceeded" };

/** Stage section deltas after document operations, against the accepted projection. */
export const stageSectionBoundaryProperties = ({
  state,
  target,
  originalRevisionIdSeed,
  revisionStamp,
  author,
  maxRanges,
  mapTargetProperties,
}: {
  state: import("prosemirror-state").EditorState;
  target: PMNode;
  originalRevisionIdSeed: number;
  revisionStamp: { idSeed: number; date: string };
  author: string;
  maxRanges: number;
  mapTargetProperties: (args: {
    kind: "inserted" | "retained";
    current: SectionProperties | undefined;
    target: SectionProperties;
  }) => MappedSectionBoundaryProperties | null;
}): StageSectionBoundaryPropertiesResult => {
  if (!Number.isSafeInteger(maxRanges) || maxRanges < 0) return { status: "budget-exceeded" };
  const reviewedState = resolveAllChangesInHeadlessStateWithMapping(state, "accept");
  const reviewed = boundaryParagraphEntriesOf(reviewedState.state.doc);
  const targets = boundaryParagraphEntriesOf(target);
  if (!reviewed || !targets) {
    return { status: "unalignable", detail: "section boundary revision history is unsupported" };
  }
  if (reviewed.length !== targets.length) {
    return { status: "unalignable", detail: "retained paragraph count differs" };
  }
  type Update = { position: number; target: SectionProperties; previous?: SectionProperties };
  const updates: Update[] = [];
  const inverse = reviewedState.mapping.invert();
  for (const [index, reviewedParagraph] of reviewed.entries()) {
    const targetParagraph = targets[index];
    if (!targetParagraph || reviewedParagraph.text !== targetParagraph.text) {
      return { status: "unalignable", detail: `retained paragraph text differs at ${index}` };
    }
    if (targetParagraph.sectionProperties === undefined) {
      if (reviewedParagraph.sectionProperties !== undefined) {
        return { status: "unalignable", detail: "section endpoint presence differs" };
      }
      continue;
    }
    const targetProperties = withoutChangeHistory(targetParagraph.sectionProperties);
    if (
      reviewedParagraph.sectionProperties !== undefined &&
      canonicalJson(withoutChangeHistory(reviewedParagraph.sectionProperties)) ===
        canonicalJson(targetProperties)
    ) {
      continue;
    }
    const mapped = inverse.mapResult(reviewedParagraph.position, 1);
    if (mapped.deleted) return { status: "unalignable", detail: "boundary has no source position" };
    const current = state.doc.nodeAt(mapped.pos);
    if (!current || current.type.name !== "paragraph") {
      return { status: "unalignable", detail: "mapped boundary is not a paragraph" };
    }
    const attrs = expectParagraphAttrs(current);
    const currentProperties = attrs._sectionProperties;
    const marker = attrs.pPrMark;
    if (marker?.kind === "ins" && marker.info.id >= originalRevisionIdSeed) {
      if (currentProperties !== undefined) {
        return {
          status: "unalignable",
          detail: "inserted endpoint already has section properties",
        };
      }
      const staged = mapTargetProperties({
        kind: "inserted",
        current: undefined,
        target: targetProperties,
      });
      if (!staged || staged.kind !== "inserted") {
        return { status: "unalignable", detail: "inserted section references cannot be mapped" };
      }
      updates.push({ position: mapped.pos, target: staged.target });
      continue;
    }
    if (currentProperties === undefined || currentProperties.propertyChanges !== undefined) {
      return { status: "unalignable", detail: "retained endpoint cannot carry a section change" };
    }
    const staged = mapTargetProperties({
      kind: "retained",
      current: withoutChangeHistory(currentProperties),
      target: targetProperties,
    });
    if (!staged || staged.kind !== "retained") {
      return { status: "unalignable", detail: "retained section references cannot be mapped" };
    }
    updates.push({ position: mapped.pos, previous: staged.previous, target: staged.target });
  }
  if (updates.length > maxRanges) return { status: "budget-exceeded" };
  let nextRevisionId = revisionStamp.idSeed;
  let transaction = state.tr;
  for (const update of updates.toSorted((left, right) => right.position - left.position)) {
    const position = transaction.mapping.map(update.position, 1);
    const node = transaction.doc.nodeAt(position);
    if (!node || node.type.name !== "paragraph") {
      return { status: "unalignable", detail: "boundary moved while staging" };
    }
    const previousReferences =
      update.previous === undefined
        ? undefined
        : sectionReferenceHistory({ previous: update.previous, target: update.target });
    const propertyChanges =
      update.previous === undefined
        ? undefined
        : [
            {
              type: "sectionPropertyChange" as const,
              info: { id: nextRevisionId++, author, date: revisionStamp.date },
              previousProperties: update.previous,
              ...(previousReferences !== undefined && { previousReferences }),
              currentProperties: update.target,
            },
          ];
    transaction = transaction.setNodeMarkup(position, undefined, {
      ...node.attrs,
      _sectionProperties: {
        ...update.target,
        ...(propertyChanges === undefined ? {} : { propertyChanges }),
      },
      sectionBreakType: update.target.sectionStart ?? null,
    });
  }
  return { status: "matched", transaction, nextRevisionId, rangeCount: updates.length };
};
