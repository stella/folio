import type { Node as PMNode } from "prosemirror-model";

import {
  DOCUMENT_OP_TYPES,
  type SetSectionEndpointOp,
} from "@stll/docx-core/ops";

import { sectionReferenceHistory } from "../docx/sectionReferenceHistory";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import { resolveAllChangesInHeadlessStateWithMapping } from "../prosemirror/commands/comments";
import type {
  Document,
  Paragraph,
  ParagraphMarkChange,
  SectionProperties,
} from "../types/document";
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
  paraId: string | undefined;
  text: string;
  sectionProperties: SectionProperties | undefined;
};

type BoundaryChild =
  | { type: "paragraph"; paragraph: BoundaryParagraph }
  | { type: "other"; nodeType: string };

const withoutChangeHistory = (
  properties: SectionProperties,
): SectionProperties => {
  const { propertyChanges: _propertyChanges, ...result } = properties;
  return result;
};

const boundaryChildrenOf = (document: PMNode): BoundaryChild[] | null => {
  const children: BoundaryChild[] = [];
  let hasChangeHistory = false;
  document.forEach((node, position) => {
    if (node.type.name !== "paragraph") {
      // Neither an opaque block nor a block bookmark has a paragraph mark, so
      // neither can carry a section endpoint. Their preservation is verified
      // independently; they do not belong in this topology.
      if (
        node.type.name === "preservedBlock" ||
        node.type.name === "blockBookmarkBoundary"
      )
        return;
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
      paragraph: {
        position,
        paraId: attrs.paraId ?? undefined,
        text: node.textContent,
        sectionProperties,
      },
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
  mapTargetProperties?: (
    properties: SectionProperties,
  ) => SectionProperties | null;
}): SectionBoundaryPropertiesComparison => {
  const currentChildren = boundaryChildrenOf(current);
  const targetChildren = boundaryChildrenOf(target);
  if (!currentChildren || !targetChildren) {
    return {
      status: "unalignable",
      detail: "section boundary revision history is unsupported",
    };
  }
  if (currentChildren.length !== targetChildren.length) {
    return {
      status: "unalignable",
      detail: "section boundary topology differs",
    };
  }

  const changes: SectionBoundaryPropertyChange[] = [];
  for (const [index, currentChild] of currentChildren.entries()) {
    const targetChild = targetChildren[index];
    if (!targetChild || currentChild.type !== targetChild.type) {
      return {
        status: "unalignable",
        detail: "section boundary topology differs",
      };
    }
    if (currentChild.type === "other" && targetChild.type === "other") {
      if (currentChild.nodeType !== targetChild.nodeType) {
        return {
          status: "unalignable",
          detail: "section boundary topology differs",
        };
      }
      continue;
    }
    if (currentChild.type !== "paragraph" || targetChild.type !== "paragraph") {
      return {
        status: "unalignable",
        detail: "section boundary topology differs",
      };
    }
    const currentParagraph = currentChild.paragraph;
    const targetParagraph = targetChild.paragraph;
    if (currentParagraph.text !== targetParagraph.text) {
      return {
        status: "unalignable",
        detail: "retained paragraph text differs",
      };
    }
    const currentProperties = currentParagraph.sectionProperties;
    const targetProperties = targetParagraph.sectionProperties;
    if (currentProperties === undefined || targetProperties === undefined) {
      if (currentProperties !== targetProperties) {
        return {
          status: "unalignable",
          detail: "section endpoint presence differs",
        };
      }
      continue;
    }
    const currentWithoutHistory = withoutChangeHistory(currentProperties);
    const mappedTarget = mapTargetProperties
      ? mapTargetProperties(targetProperties)
      : targetProperties;
    if (mappedTarget === null)
      return {
        status: "unalignable",
        detail: "section references cannot be mapped",
      };
    const targetWithoutHistory = withoutChangeHistory(mappedTarget);
    if (
      canonicalJson(currentWithoutHistory) ===
      canonicalJson(targetWithoutHistory)
    )
      continue;
    changes.push({
      position: currentParagraph.position,
      current: currentWithoutHistory,
      target: targetWithoutHistory,
    });
  }
  return { status: "matched", changes };
};

const boundaryParagraphEntriesOf = (
  document: PMNode,
): BoundaryParagraph[] | null => {
  const children = boundaryChildrenOf(document);
  if (!children) return null;
  return children.flatMap((child) =>
    child.type === "paragraph" ? [child.paragraph] : [],
  );
};

const hasSectionEndpoint = (document: PMNode): boolean => {
  let found = false;
  document.forEach((node) => {
    if (
      node.type.name === "paragraph" &&
      expectParagraphAttrs(node)._sectionProperties !== undefined
    ) {
      found = true;
    }
  });
  return found;
};

type MappedSectionBoundaryProperties =
  | { kind: "inserted"; target: SectionProperties }
  | {
      kind: "retained";
      previous: SectionProperties;
      target: SectionProperties;
    };

type MapTargetProperties = (args: {
  kind: "inserted" | "retained";
  current: SectionProperties | undefined;
  target: SectionProperties;
}) => MappedSectionBoundaryProperties | null;

/**
 * Where an accepted boundary paragraph stands in the document being staged:
 * its key there, its section properties and its paragraph mark, or why it
 * has no such place.
 */
type BoundarySource<Key> =
  | {
      found: true;
      key: Key;
      sectionProperties: SectionProperties | undefined;
      pPrMark: ParagraphMarkChange | undefined;
    }
  | { found: false; detail: string };

/** One section endpoint to stage, in document order. */
type BoundaryUpdate<Key> = {
  key: Key;
  target: SectionProperties;
  previous?: SectionProperties;
};

type BoundaryPlan<Key> =
  | { status: "matched"; updates: readonly BoundaryUpdate<Key>[] }
  | { status: "unalignable"; detail: string }
  | { status: "budget-exceeded" };

/**
 * The section endpoints a target needs, planned on the accepted projection
 * of the document being staged: each accepted boundary paragraph is found
 * where it stands through `sourceOf`.
 */
const planSectionBoundaryUpdates = <Key>({
  accepted,
  target,
  originalRevisionIdSeed,
  maxRanges,
  mapTargetProperties,
  sourceOf,
}: {
  accepted: PMNode;
  target: PMNode;
  originalRevisionIdSeed: number;
  maxRanges: number;
  mapTargetProperties: MapTargetProperties;
  sourceOf: (paragraph: BoundaryParagraph) => BoundarySource<Key>;
}): BoundaryPlan<Key> => {
  const reviewed = boundaryParagraphEntriesOf(accepted);
  const targets = boundaryParagraphEntriesOf(target);
  if (!reviewed || !targets) {
    return {
      status: "unalignable",
      detail: "section boundary revision history is unsupported",
    };
  }
  if (reviewed.length !== targets.length) {
    return {
      status: "unalignable",
      detail: "retained paragraph count differs",
    };
  }
  const updates: BoundaryUpdate<Key>[] = [];
  for (const [index, reviewedParagraph] of reviewed.entries()) {
    const targetParagraph = targets[index];
    if (!targetParagraph || reviewedParagraph.text !== targetParagraph.text) {
      return {
        status: "unalignable",
        detail: `retained paragraph text differs at ${index}`,
      };
    }
    if (targetParagraph.sectionProperties === undefined) {
      if (reviewedParagraph.sectionProperties !== undefined) {
        return {
          status: "unalignable",
          detail: "section endpoint presence differs",
        };
      }
      continue;
    }
    const targetProperties = withoutChangeHistory(
      targetParagraph.sectionProperties,
    );
    if (
      reviewedParagraph.sectionProperties !== undefined &&
      canonicalJson(
        withoutChangeHistory(reviewedParagraph.sectionProperties),
      ) === canonicalJson(targetProperties)
    ) {
      continue;
    }
    const source = sourceOf(reviewedParagraph);
    if (!source.found) return { status: "unalignable", detail: source.detail };
    const currentProperties = source.sectionProperties;
    const marker = source.pPrMark;
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
        return {
          status: "unalignable",
          detail: "inserted section references cannot be mapped",
        };
      }
      updates.push({ key: source.key, target: staged.target });
      continue;
    }
    if (
      currentProperties === undefined ||
      currentProperties.propertyChanges !== undefined
    ) {
      return {
        status: "unalignable",
        detail: "retained endpoint cannot carry a section change",
      };
    }
    const staged = mapTargetProperties({
      kind: "retained",
      current: withoutChangeHistory(currentProperties),
      target: targetProperties,
    });
    if (!staged || staged.kind !== "retained") {
      return {
        status: "unalignable",
        detail: "retained section references cannot be mapped",
      };
    }
    updates.push({
      key: source.key,
      previous: staged.previous,
      target: staged.target,
    });
  }
  if (updates.length > maxRanges) return { status: "budget-exceeded" };
  return { status: "matched", updates };
};

/**
 * An endpoint's section properties with the tracked change that stages
 * `update` on it, under revision `id`.
 */
const stagedBoundaryProperties = (
  update: BoundaryUpdate<unknown>,
  id: number,
  author: string,
  date: string,
): SectionProperties => {
  const previousReferences =
    update.previous === undefined
      ? {}
      : sectionReferenceHistory({
          previous: update.previous,
          target: update.target,
        });
  return {
    ...update.target,
    propertyChanges: [
      {
        type: "sectionPropertyChange" as const,
        info: { id, author, date },
        ...(update.previous === undefined
          ? {}
          : { previousProperties: update.previous }),
        ...(previousReferences !== undefined && { previousReferences }),
        currentProperties: update.target,
      },
    ],
  };
};

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
  mapTargetProperties: MapTargetProperties;
}): StageSectionBoundaryPropertiesResult => {
  if (!Number.isSafeInteger(maxRanges) || maxRanges < 0)
    return { status: "budget-exceeded" };
  // A target no paragraph of which ends a section has no endpoint to stage,
  // whatever the accepted projection says, so it is not worth building: that
  // is an accept-all over the whole redlined story. Whether the rest of the
  // accepted story agrees with the target is the round-trip check's question.
  if (!hasSectionEndpoint(target)) {
    return {
      status: "matched",
      transaction: state.tr,
      nextRevisionId: revisionStamp.idSeed,
      rangeCount: 0,
    };
  }
  const reviewedState = resolveAllChangesInHeadlessStateWithMapping(
    state,
    "accept",
  );
  const inverse = reviewedState.mapping.invert();
  const plan = planSectionBoundaryUpdates<number>({
    accepted: reviewedState.state.doc,
    target,
    originalRevisionIdSeed,
    maxRanges,
    mapTargetProperties,
    sourceOf: (paragraph) => {
      const mapped = inverse.mapResult(paragraph.position, 1);
      if (mapped.deleted)
        return { found: false, detail: "boundary has no source position" };
      const current = state.doc.nodeAt(mapped.pos);
      if (!current || current.type.name !== "paragraph")
        return { found: false, detail: "mapped boundary is not a paragraph" };
      const attrs = expectParagraphAttrs(current);
      return {
        found: true,
        key: mapped.pos,
        sectionProperties: attrs._sectionProperties,
        pPrMark: attrs.pPrMark ?? undefined,
      };
    },
  });
  if (plan.status !== "matched") return plan;
  let nextRevisionId = revisionStamp.idSeed;
  let transaction = state.tr;
  for (const update of plan.updates.toSorted(
    (left, right) => right.key - left.key,
  )) {
    const position = transaction.mapping.map(update.key, 1);
    const node = transaction.doc.nodeAt(position);
    if (!node || node.type.name !== "paragraph") {
      return { status: "unalignable", detail: "boundary moved while staging" };
    }
    transaction = transaction.setNodeMarkup(position, undefined, {
      ...node.attrs,
      _sectionProperties: stagedBoundaryProperties(
        update,
        nextRevisionId++,
        author,
        revisionStamp.date,
      ),
    });
  }
  return {
    status: "matched",
    transaction,
    nextRevisionId,
    rangeCount: plan.updates.length,
  };
};

export type StageSectionBoundaryOperationsResult =
  | {
      status: "matched";
      operations: readonly SetSectionEndpointOp[];
      nextRevisionId: number;
    }
  | { status: "unalignable"; detail: string }
  | { status: "budget-exceeded" };

/**
 * Stage section deltas on a document model: planned as
 * {@link stageSectionBoundaryProperties} plans them, on `accepted`, the
 * accepted projection of `document`, with each boundary paragraph found in
 * `document` by its paragraph id. One section-endpoint operation per
 * endpoint, revision ids given from the last endpoint back as the
 * transaction gives them.
 */
export const stageSectionBoundaryOperations = ({
  document,
  accepted,
  target,
  originalRevisionIdSeed,
  revisionStamp,
  author,
  maxRanges,
  mapTargetProperties,
}: {
  document: Document;
  accepted: () => PMNode;
  target: PMNode;
  originalRevisionIdSeed: number;
  revisionStamp: { idSeed: number; date: string };
  author: string;
  maxRanges: number;
  mapTargetProperties: MapTargetProperties;
}): StageSectionBoundaryOperationsResult => {
  if (!Number.isSafeInteger(maxRanges) || maxRanges < 0)
    return { status: "budget-exceeded" };
  if (!hasSectionEndpoint(target))
    return {
      status: "matched",
      operations: [],
      nextRevisionId: revisionStamp.idSeed,
    };
  // Section endpoints are top-level paragraphs of the body.
  const paragraphs = new Map<string, Paragraph[]>();
  for (const block of document.package.document.content) {
    if (block.type !== "paragraph" || block.paraId === undefined) continue;
    const key = block.paraId.toUpperCase();
    paragraphs.set(key, [...(paragraphs.get(key) ?? []), block]);
  }
  const plan = planSectionBoundaryUpdates<{ paraId: string; order: number }>({
    accepted: accepted(),
    target,
    originalRevisionIdSeed,
    maxRanges,
    mapTargetProperties,
    sourceOf: (paragraph) => {
      const found =
        paragraph.paraId === undefined
          ? []
          : (paragraphs.get(paragraph.paraId.toUpperCase()) ?? []);
      const [current] = found;
      if (current === undefined || current.paraId === undefined)
        return { found: false, detail: "boundary has no source position" };
      if (found.length > 1)
        return { found: false, detail: "boundary paragraph id is ambiguous" };
      return {
        found: true,
        key: { paraId: current.paraId, order: paragraph.position },
        sectionProperties: current.sectionProperties,
        pPrMark: current.pPrMark,
      };
    },
  });
  if (plan.status !== "matched") return plan;
  let nextRevisionId = revisionStamp.idSeed;
  const operations = plan.updates
    .toSorted((left, right) => right.key.order - left.key.order)
    .map((update): SetSectionEndpointOp => ({
      type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
      endpoint: { type: "paragraph", blockId: update.key.paraId },
      properties: {
        type: "present",
        value: stagedBoundaryProperties(
          update,
          nextRevisionId++,
          author,
          revisionStamp.date,
        ),
      },
    }));
  return { status: "matched", operations, nextRevisionId };
};
