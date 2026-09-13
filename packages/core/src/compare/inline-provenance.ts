import type { Mark, Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";

import type { FolioRevisionStamp } from "../ai-edits/apply";
import { sourceDocumentOf, styleResolverOf } from "../ai-edits/snapshot";
import type { FolioAIEditSnapshot } from "../ai-edits/types";
import { expectRunPropertyChangeMarkAttrs, expectTableCellAttrs } from "../prosemirror/attrs";
import {
  applyMarksToRunFormattingRepresentation,
  expandRunFormattingCarrier,
  runFormattingCarrierReviewText,
  type RunFormattingCarrier,
  type RunFormattingCarrierRepresentation,
} from "../prosemirror/runFormattingInlineCarriers";
import { getDocumentStyleResolver } from "../prosemirror/plugins/documentStyles";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "../prosemirror/runFormattingReconciliation";
import {
  paragraphRunStyleContextAt,
  type RunStyleResolver,
} from "../prosemirror/runStyleFormatting";
import { isTableCellRetainedInReviewView } from "../prosemirror/tableCellRevisionVisibility";
import type { RunPropertyChange, TextFormatting } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";

export type InlineProvenanceTargetOptions = {
  targetSnapshot: FolioAIEditSnapshot;
  revisionStamp: FolioRevisionStamp;
  /** First id reserved for this comparison, before other planned operations allocated ids. */
  originalRevisionIdSeed: number;
  maxRanges: number;
};

export type MatchInlineProvenanceOptions = InlineProvenanceTargetOptions & {
  state: EditorState;
  author: string;
};

export type MatchInlineProvenanceResult =
  | {
      status: "matched";
      transaction: Transaction;
      nextRevisionId: number;
      changedTargetBlockIds: readonly string[];
      rangeCount: number;
    }
  | { status: "unalignable" }
  | { status: "budget-exceeded" };

type InlineCarrier = {
  carrier: RunFormattingCarrier;
  text: string;
  targetBlockId: string | undefined;
};

type CarrierRepresentation = RunFormattingCarrierRepresentation & {
  from: number;
  to: number;
};

type MatchedRepresentation = {
  live: CarrierRepresentation;
  target: CarrierRepresentation;
  targetBlockId: string;
};

type PlannedChange = MatchedRepresentation & {
  formatting: TextFormatting;
  previousFormatting: TextFormatting;
  inserted: boolean;
  propertyChange: Mark | undefined;
};

const isDeleted = (carrier: RunFormattingCarrier): boolean =>
  carrier.representations.some(({ node }) =>
    node.marks.some(({ type }) => type.name === "deletion"),
  );

const isInserted = (node: PMNode): boolean =>
  node.marks.some(({ type }) => type.name === "insertion");

const carrierRepresentations = (carrier: RunFormattingCarrier): CarrierRepresentation[] =>
  carrier.representations.map((representation) => ({
    ...representation,
    from: representation.position,
    to: representation.position + representation.node.nodeSize,
  }));

const targetBlockIdLookup = (anchors: FolioAIEditSnapshot["anchors"]) => {
  const ordered = Object.values(anchors).toSorted(
    (left, right) => left.from - right.from || left.to - right.to,
  );
  const active: (typeof ordered)[number][] = [];
  let nextAnchor = 0;

  return ({ position, nodeSize }: { position: number; nodeSize: number }): string | undefined => {
    while (nextAnchor < ordered.length) {
      const anchor = ordered.at(nextAnchor);
      if (!anchor || anchor.from > position) break;
      active.push(anchor);
      nextAnchor++;
    }
    const end = position + nodeSize;
    let owner: (typeof ordered)[number] | undefined;
    for (let index = active.length - 1; index >= 0; index--) {
      const anchor = active[index];
      if (!anchor) {
        continue;
      }
      if (anchor.to < end) {
        active.splice(index, 1);
        continue;
      }
      if (
        anchor.from <= position &&
        end <= anchor.to &&
        (owner === undefined || anchor.to - anchor.from < owner.to - owner.from)
      ) {
        owner = anchor;
      }
    }
    return owner?.id;
  };
};

const collectCarriers = ({
  doc,
  targetSnapshot,
}: {
  doc: PMNode;
  targetSnapshot?: FolioAIEditSnapshot;
}): InlineCarrier[] | null => {
  const carriers: InlineCarrier[] = [];
  const targetBlockIdAt = targetSnapshot ? targetBlockIdLookup(targetSnapshot.anchors) : undefined;
  let unanchoredTargetCarrier = false;
  doc.descendants((node, position) => {
    if (
      (node.type.name === "tableCell" || node.type.name === "tableHeader") &&
      !isTableCellRetainedInReviewView(expectTableCellAttrs(node).cellMarker, "final")
    ) {
      return false;
    }
    if (!node.isInline) {
      return true;
    }
    const carrier = expandRunFormattingCarrier(node, position);
    if (!carrier) {
      return !node.isAtom;
    }
    if (isDeleted(carrier)) {
      return false;
    }
    const targetBlockId = targetBlockIdAt?.({ position, nodeSize: node.nodeSize });
    if (targetSnapshot && targetBlockId === undefined) {
      unanchoredTargetCarrier = true;
      return false;
    }
    carriers.push({ carrier, text: runFormattingCarrierReviewText(carrier), targetBlockId });
    return false;
  });
  return unanchoredTargetCarrier ? null : carriers;
};

const sameCarrierShape = (left: InlineCarrier, right: InlineCarrier): boolean => {
  if (left.carrier.disposition !== right.carrier.disposition || left.text !== right.text) {
    return false;
  }
  const leftRepresentations = left.carrier.representations;
  const rightRepresentations = right.carrier.representations;
  return (
    leftRepresentations.length === rightRepresentations.length &&
    leftRepresentations.every((representation, index) => {
      const target = rightRepresentations.at(index);
      return (
        target !== undefined &&
        representation.role === target.role &&
        representation.node.type === target.node.type
      );
    })
  );
};

/**
 * Match whole formatting carriers while permitting only text runs to be cut at
 * UTF-16 boundaries. Non-text carriers are serialization atoms and therefore
 * must line up exactly on both sides.
 */
const matchCarrierStreams = ({
  live,
  target,
}: {
  live: readonly InlineCarrier[];
  target: readonly InlineCarrier[];
}): MatchedRepresentation[] | null => {
  const matched: MatchedRepresentation[] = [];
  let liveIndex = 0;
  let targetIndex = 0;
  let liveOffset = 0;
  let targetOffset = 0;

  while (liveIndex < live.length || targetIndex < target.length) {
    const liveCarrier = live.at(liveIndex);
    const targetCarrier = target.at(targetIndex);
    if (!liveCarrier || !targetCarrier || targetCarrier.targetBlockId === undefined) {
      return null;
    }
    const liveText = liveCarrier.text.slice(liveOffset);
    const targetText = targetCarrier.text.slice(targetOffset);
    const liveIsText = liveCarrier.carrier.disposition === "text-run";
    const targetIsText = targetCarrier.carrier.disposition === "text-run";

    if (!liveIsText || !targetIsText) {
      if (liveOffset !== 0 || targetOffset !== 0 || !sameCarrierShape(liveCarrier, targetCarrier)) {
        return null;
      }
      const liveRepresentations = carrierRepresentations(liveCarrier.carrier);
      const targetRepresentations = carrierRepresentations(targetCarrier.carrier);
      for (const [index, liveRepresentation] of liveRepresentations.entries()) {
        const targetRepresentation = targetRepresentations.at(index);
        if (!targetRepresentation) {
          return null;
        }
        matched.push({
          live: liveRepresentation,
          target: targetRepresentation,
          targetBlockId: targetCarrier.targetBlockId,
        });
      }
      liveIndex++;
      targetIndex++;
      continue;
    }

    const sharedLength = Math.min(liveText.length, targetText.length);
    if (
      sharedLength === 0 ||
      liveText.slice(0, sharedLength) !== targetText.slice(0, sharedLength)
    ) {
      return null;
    }
    const liveRepresentation = carrierRepresentations(liveCarrier.carrier).at(0);
    const targetRepresentation = carrierRepresentations(targetCarrier.carrier).at(0);
    if (!liveRepresentation || !targetRepresentation) {
      return null;
    }
    matched.push({
      live: {
        ...liveRepresentation,
        from: liveRepresentation.from + liveOffset,
        to: liveRepresentation.from + liveOffset + sharedLength,
      },
      target: {
        ...targetRepresentation,
        from: targetRepresentation.from + targetOffset,
        to: targetRepresentation.from + targetOffset + sharedLength,
      },
      targetBlockId: targetCarrier.targetBlockId,
    });
    liveOffset += sharedLength;
    targetOffset += sharedLength;
    if (liveOffset === liveCarrier.text.length) {
      liveIndex++;
      liveOffset = 0;
    }
    if (targetOffset === targetCarrier.text.length) {
      targetIndex++;
      targetOffset = 0;
    }
  }
  return matched;
};

const authoredFormattingAt = ({
  doc,
  representation,
  styleResolver,
}: {
  doc: PMNode;
  representation: CarrierRepresentation;
  styleResolver: RunStyleResolver | null;
}): TextFormatting =>
  readAuthoredRunFormatting({
    context: paragraphRunStyleContextAt({ doc, pos: representation.from, styleResolver }),
    marks: representation.node.marks,
    styleResolver,
  });

const changeWithCurrentFormatting = ({
  change,
  formatting,
}: {
  change: RunPropertyChange;
  formatting: TextFormatting;
}): RunPropertyChange => ({ ...change, currentFormatting: formatting });

/**
 * Reconcile direct inline run formatting after the comparison plan has aligned
 * text. It deliberately performs all stream, revision-slot, and range-budget
 * checks before creating a transaction step.
 */
export const matchInlineProvenance = ({
  state,
  targetSnapshot,
  revisionStamp,
  originalRevisionIdSeed,
  author,
  maxRanges,
}: MatchInlineProvenanceOptions): MatchInlineProvenanceResult => {
  const targetStyleResolver = styleResolverOf(targetSnapshot);
  if (!Number.isSafeInteger(maxRanges) || maxRanges < 0) {
    return { status: "budget-exceeded" };
  }
  const targetDocument = sourceDocumentOf(targetSnapshot);
  const liveCarriers = collectCarriers({ doc: state.doc });
  const targetCarriers = collectCarriers({ doc: targetDocument, targetSnapshot });
  if (!liveCarriers || !targetCarriers) {
    return { status: "unalignable" };
  }
  const matched = matchCarrierStreams({ live: liveCarriers, target: targetCarriers });
  if (!matched) {
    return { status: "unalignable" };
  }

  const propertyChangeType = state.schema.marks["runPropertyChange"];
  const liveStyleResolver = getDocumentStyleResolver(state);
  const planned: PlannedChange[] = [];
  for (const segment of matched) {
    const liveFormatting = authoredFormattingAt({
      doc: state.doc,
      representation: segment.live,
      styleResolver: liveStyleResolver,
    });
    const targetFormatting = authoredFormattingAt({
      doc: targetDocument,
      representation: segment.target,
      styleResolver: targetStyleResolver,
    });
    if (canonicalJson(liveFormatting) === canonicalJson(targetFormatting)) {
      continue;
    }
    const existingPropertyChange = propertyChangeType
      ? segment.live.node.marks.find((mark) => mark.type === propertyChangeType)
      : undefined;
    const propertyChange =
      existingPropertyChange &&
      expectRunPropertyChangeMarkAttrs(existingPropertyChange).changes.length > 0
        ? existingPropertyChange
        : undefined;
    if (
      propertyChange &&
      expectRunPropertyChangeMarkAttrs(propertyChange).changes.some(
        ({ info }) => info.id < originalRevisionIdSeed,
      )
    ) {
      return { status: "unalignable" };
    }
    if (!isInserted(segment.live.node) && !propertyChangeType) {
      return { status: "unalignable" };
    }
    planned.push({
      ...segment,
      formatting: targetFormatting,
      previousFormatting: liveFormatting,
      inserted: isInserted(segment.live.node),
      propertyChange,
    });
  }
  if (planned.length > maxRanges) {
    return { status: "budget-exceeded" };
  }

  const transaction = state.tr;
  let nextRevisionId = revisionStamp.idSeed;
  const changedTargetBlockIds = new Set<string>();
  for (const change of planned) {
    const context = paragraphRunStyleContextAt({
      doc: state.doc,
      pos: change.live.from,
      styleResolver: liveStyleResolver,
    });
    const formattingMarks = reconcileRunFormattingMarks({
      authoredFormatting: change.formatting,
      context,
      node: change.live.node,
      styleResolver: liveStyleResolver,
    });
    let marks = formattingMarks;
    if (!change.inserted && propertyChangeType) {
      if (change.propertyChange) {
        const attrs = expectRunPropertyChangeMarkAttrs(change.propertyChange);
        marks = propertyChangeType
          .create({
            ...change.propertyChange.attrs,
            changes: attrs.changes.map((entry) =>
              changeWithCurrentFormatting({ change: entry, formatting: change.formatting }),
            ),
          })
          .addToSet(formattingMarks);
      } else {
        const propertyChange: RunPropertyChange = {
          type: "runPropertyChange",
          info: { id: nextRevisionId++, author, date: revisionStamp.date },
          ...(Object.keys(change.previousFormatting).length > 0
            ? { previousFormatting: change.previousFormatting }
            : {}),
        };
        marks = propertyChangeType.create({ changes: [propertyChange] }).addToSet(formattingMarks);
      }
    }
    applyMarksToRunFormattingRepresentation({
      tr: transaction,
      representation: change.live,
      marks,
    });
    changedTargetBlockIds.add(change.targetBlockId);
  }
  return {
    status: "matched",
    transaction,
    nextRevisionId,
    changedTargetBlockIds: [...changedTargetBlockIds],
    rangeCount: planned.length,
  };
};

/** Compare full authored run-property provenance in two snapshot source documents. */
export const sameAuthoredInlineProvenance = (
  baseSnapshot: FolioAIEditSnapshot,
  targetSnapshot: FolioAIEditSnapshot,
): boolean => {
  const baseDocument = sourceDocumentOf(baseSnapshot);
  const targetDocument = sourceDocumentOf(targetSnapshot);
  const baseCarriers = collectCarriers({ doc: baseDocument, targetSnapshot: baseSnapshot });
  const targetCarriers = collectCarriers({ doc: targetDocument, targetSnapshot });
  if (!baseCarriers || !targetCarriers) {
    return false;
  }
  const matched = matchCarrierStreams({ live: baseCarriers, target: targetCarriers });
  if (!matched) {
    return false;
  }
  const baseStyleResolver = styleResolverOf(baseSnapshot);
  const targetStyleResolver = styleResolverOf(targetSnapshot);
  return matched.every(
    ({ live, target }) =>
      canonicalJson(
        authoredFormattingAt({
          doc: baseDocument,
          representation: live,
          styleResolver: baseStyleResolver,
        }),
      ) ===
      canonicalJson(
        authoredFormattingAt({
          doc: targetDocument,
          representation: target,
          styleResolver: targetStyleResolver,
        }),
      ),
  );
};
