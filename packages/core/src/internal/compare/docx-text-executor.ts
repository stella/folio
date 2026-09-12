import { Mark, type MarkType, type Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { panic } from "better-result";
import { sameTextFormatting } from "@stll/docx-core/model";

import { buildCleanBlockText } from "../../ai-edits/clean-text";
import {
  expectHyperlinkMarkAttrs,
  expectRunPropertyChangeMarkAttrs,
} from "../../prosemirror/attrs";
import { getDocumentStyleResolver } from "../../prosemirror/plugins/documentStyles";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "../../prosemirror/runFormattingReconciliation";
import {
  selectRunFormattingCarrierRepresentations,
  type SelectedRunFormattingCarrierRepresentation,
} from "../../prosemirror/runFormattingInlineCarriers";
import { paragraphRunStyleContextAt } from "../../prosemirror/runStyleFormatting";
import type { RunPropertyChange, TextFormatting } from "../../types/document";
import type {
  DocxComparisonDeletedFragment,
  DocxComparisonEqualFragment,
  DocxComparisonRangePlan,
  DocxInlineContainer,
} from "./docx-program";
import { docxInlineOccurrenceKey } from "./docx-program";

type ExactFormattingMetadata = {
  readonly author: string;
  readonly date: string;
};

const applyMarksToRunFormattingRepresentation = ({
  tr,
  representation,
  marks,
}: {
  readonly tr: Transaction;
  readonly representation: SelectedRunFormattingCarrierRepresentation;
  readonly marks: readonly Mark[];
}): void => {
  const { node, position, from, to } = representation;
  if (Mark.sameSet(node.marks, marks)) return;
  if (!node.isText) {
    tr.setNodeMarkup(position, undefined, node.attrs, marks);
    return;
  }
  for (const current of node.marks) {
    if (!marks.some((candidate) => candidate.eq(current))) {
      tr.removeMark(from, to, current.type);
    }
  }
  for (const next of marks) {
    if (!node.marks.some((candidate) => candidate.eq(next))) {
      tr.addMark(from, to, next);
    }
  }
};

type ExactRangePosition = {
  readonly type: "format";
  readonly from: number;
  readonly to: number;
  readonly fragment: DocxComparisonEqualFragment;
};

type ExactReplacementStep =
  | {
      readonly type: "del";
      readonly from: number;
      readonly to: number;
      readonly fragment: DocxComparisonDeletedFragment;
    }
  | {
      readonly type: "ins";
      readonly at: number;
      readonly text: string;
      readonly formatting: Readonly<TextFormatting>;
      readonly targetInlineContainers: readonly DocxInlineContainer[];
    }
  | ExactRangePosition;

type TargetInlineReuse = {
  readonly occurrenceKey: string;
  readonly mark: Mark;
};

type ExactReplacementPlan = {
  readonly steps: readonly ExactReplacementStep[];
  readonly retained: readonly {
    readonly from: number;
    readonly to: number;
    readonly targetInlineContainers: readonly DocxInlineContainer[];
  }[];
};

export type DocxTextRangeTarget = {
  readonly blockNode: PMNode;
  readonly blockFrom: number;
  readonly sourceStartOffset: number;
  readonly range: DocxComparisonRangePlan;
  readonly styleResolver: ReturnType<typeof getDocumentStyleResolver>;
};

export type DocxTextRangePreflight =
  | {
      readonly type: "ready";
      readonly steps: readonly ExactReplacementStep[];
      readonly targetInlineReuse: readonly TargetInlineReuse[];
    }
  | {
      readonly type: "unsupported";
      readonly reason: "source-mismatch" | "source-formatting-mismatch" | "pending-run-change";
    };

const hasPendingRunPropertyChange = (
  doc: PMNode,
  propertyChangeType: MarkType | undefined,
  from: number,
  to: number,
): boolean =>
  propertyChangeType !== undefined &&
  selectRunFormattingCarrierRepresentations({ doc, from, to }).some(({ node }) => {
    const existing = node.marks.find((mark) => mark.type === propertyChangeType);
    return existing ? expectRunPropertyChangeMarkAttrs(existing).changes.length > 0 : false;
  });

const exactReplacementSteps = ({
  blockNode,
  blockFrom,
  range,
  sourceStartOffset,
}: DocxTextRangeTarget): ExactReplacementPlan | null => {
  if (blockNode.content.size !== blockNode.textContent.length) return null;
  const clean = buildCleanBlockText(blockNode, blockFrom);
  if (
    sourceStartOffset < 0 ||
    clean.text.slice(sourceStartOffset, sourceStartOffset + range.sourceText.length) !==
      range.sourceText
  ) {
    return null;
  }
  const offsetAt = (offset: number): number | null =>
    clean.offsets[sourceStartOffset + offset] ?? null;
  const steps: ExactReplacementStep[] = [];
  const retained: ExactReplacementPlan["retained"][number][] = [];
  for (const fragment of range.fragments) {
    if (fragment.type === "ins") {
      const at = offsetAt(fragment.baseStart);
      if (at === null) return null;
      steps.push({
        type: "ins",
        at,
        text: fragment.text,
        formatting: fragment.targetFormatting,
        targetInlineContainers: fragment.targetInlineContainers,
      });
      continue;
    }
    const from = offsetAt(fragment.baseStart);
    const to = offsetAt(fragment.baseEnd);
    if (from === null || to === null) return null;
    if (fragment.type === "del") {
      steps.push({ type: "del", from, to, fragment });
      continue;
    }
    retained.push({
      from,
      to,
      targetInlineContainers: fragment.targetInlineContainers,
    });
    if (fragment.changedProperties.length > 0) {
      steps.push({ type: "format", from, to, fragment });
    }
  }
  return Object.freeze({ steps: Object.freeze(steps), retained: Object.freeze(retained) });
};

const exactTargetInlineReuse = (
  doc: PMNode,
  retained: ExactReplacementPlan["retained"],
): readonly TargetInlineReuse[] => {
  const reusableByOccurrence = new Map<string, Mark>();
  const conflictingOccurrences = new Set<string>();
  for (const { from, to, targetInlineContainers } of retained) {
    const target = targetInlineContainers.find(({ type }) => type === "hyperlink");
    if (!target) continue;
    let candidate: Mark | undefined;
    let exact = true;
    doc.nodesBetween(from, to, (node) => {
      if (!node.isInline || !exact) return;
      const hyperlink = node.marks.find(({ type }) => type.name === "hyperlink");
      if (!hyperlink) {
        exact = false;
        return;
      }
      const attrs = expectHyperlinkMarkAttrs(hyperlink);
      if (
        attrs.href !== target.href ||
        attrs.tooltip !== target.tooltip ||
        attrs.target !== target.target ||
        attrs.history !== target.history ||
        attrs.docLocation !== target.docLocation ||
        (candidate !== undefined && !candidate.eq(hyperlink))
      ) {
        exact = false;
        return;
      }
      candidate = hyperlink;
    });
    if (!exact || !candidate) continue;
    const occurrenceKey = docxInlineOccurrenceKey(target.occurrence);
    const previous = reusableByOccurrence.get(occurrenceKey);
    if (previous === undefined || previous.eq(candidate)) {
      reusableByOccurrence.set(occurrenceKey, candidate);
      continue;
    }
    reusableByOccurrence.delete(occurrenceKey);
    conflictingOccurrences.add(occurrenceKey);
  }
  return Object.freeze(
    [...reusableByOccurrence]
      .filter(([occurrenceKey]) => !conflictingOccurrences.has(occurrenceKey))
      .map(([occurrenceKey, mark]) => Object.freeze({ occurrenceKey, mark })),
  );
};

export const preflightDocxTextRange = (
  doc: PMNode,
  target: DocxTextRangeTarget,
): DocxTextRangePreflight => {
  const plan = exactReplacementSteps(target);
  if (!plan) return { type: "unsupported", reason: "source-mismatch" };
  const { steps } = plan;
  const propertyChangeType = doc.type.schema.marks["runPropertyChange"];
  if (
    steps.some(
      (step) =>
        step.type === "format" &&
        hasPendingRunPropertyChange(doc, propertyChangeType, step.from, step.to),
    )
  ) {
    return { type: "unsupported", reason: "pending-run-change" };
  }
  for (const step of steps) {
    if (step.type === "ins") continue;
    const expected = step.fragment.sourceFormatting;
    const representations = selectRunFormattingCarrierRepresentations({
      doc,
      from: step.from,
      to: step.to,
    });
    if (representations.length === 0) {
      return { type: "unsupported", reason: "source-formatting-mismatch" };
    }
    for (const representation of representations) {
      const context = paragraphRunStyleContextAt({
        doc,
        pos: representation.from,
        ...(target.styleResolver !== undefined && { styleResolver: target.styleResolver }),
      });
      const actual = readAuthoredRunFormatting({
        context,
        marks: representation.node.marks,
        ...(target.styleResolver !== undefined && { styleResolver: target.styleResolver }),
      });
      if (!sameTextFormatting(actual, expected)) {
        return { type: "unsupported", reason: "source-formatting-mismatch" };
      }
    }
  }
  return Object.freeze({
    type: "ready",
    steps,
    targetInlineReuse: exactTargetInlineReuse(doc, plan.retained),
  });
};

type ApplyDirectFormattingOptions = {
  readonly tr: Transaction;
  readonly from: number;
  readonly to: number;
  readonly formatting: Readonly<TextFormatting>;
  readonly styleResolver: ReturnType<typeof getDocumentStyleResolver>;
};

export const applyExactDirectFormatting = ({
  tr,
  from,
  to,
  formatting,
  styleResolver,
}: ApplyDirectFormattingOptions): void => {
  if (from >= to) return;
  for (const representation of selectRunFormattingCarrierRepresentations({
    doc: tr.doc,
    from,
    to,
  })) {
    const context = paragraphRunStyleContextAt({
      doc: tr.doc,
      pos: representation.from,
      ...(styleResolver !== undefined && { styleResolver }),
    });
    const marks = reconcileRunFormattingMarks({
      authoredFormatting: formatting,
      context,
      node: representation.node,
      ...(styleResolver !== undefined && { styleResolver }),
    });
    applyMarksToRunFormattingRepresentation({ tr, representation, marks });
  }
};

type TargetInlineOccurrenceAllocation = {
  readonly byOccurrence: Map<string, number>;
  next: number;
};

const targetInlineOccurrences = new WeakMap<Transaction, TargetInlineOccurrenceAllocation>();

const targetInlineOccurrence = (
  tr: Transaction,
  occurrence: DocxInlineContainer["occurrence"],
): number => {
  let allocation = targetInlineOccurrences.get(tr);
  if (!allocation) {
    let next = 0;
    tr.doc.descendants((node) => {
      const hyperlink = node.marks.find(({ type }) => type.name === "hyperlink");
      if (!hyperlink) return;
      const existing = expectHyperlinkMarkAttrs(hyperlink)._docxHyperlinkIndex;
      if (existing !== undefined) next = Math.max(next, existing + 1);
    });
    allocation = { byOccurrence: new Map(), next };
    targetInlineOccurrences.set(tr, allocation);
  }
  const occurrenceKey = docxInlineOccurrenceKey(occurrence);
  const existing = allocation.byOccurrence.get(occurrenceKey);
  if (existing !== undefined) return existing;
  const allocated = allocation.next++;
  allocation.byOccurrence.set(occurrenceKey, allocated);
  return allocated;
};

/** Recreate target-owned inline containers without copying package-local relationship ids. */
export const applyExactInlineOwnership = ({
  tr,
  from,
  to,
  containers,
  targetInlineReuse = new Map(),
}: {
  readonly tr: Transaction;
  readonly from: number;
  readonly to: number;
  readonly containers: readonly DocxInlineContainer[];
  readonly targetInlineReuse?: ReadonlyMap<string, Mark>;
}): void => {
  if (from >= to) return;
  const hyperlinkType = tr.doc.type.schema.marks["hyperlink"];
  if (!hyperlinkType) {
    return panic("A preflighted DOCX comparison lost hyperlink schema support");
  }
  // Inserted text can inherit an inclusive source mark. Clear the supported
  // container family first so the target ownership partition is exact even
  // when its branch deliberately has no owner.
  tr.removeMark(from, to, hyperlinkType);
  for (const container of containers) {
    switch (container.type) {
      case "hyperlink": {
        const reusable = targetInlineReuse.get(docxInlineOccurrenceKey(container.occurrence));
        tr.addMark(
          from,
          to,
          reusable ??
            hyperlinkType.create({
              href: container.href,
              ...(container.tooltip !== undefined && { tooltip: container.tooltip }),
              ...(container.target !== undefined && { target: container.target }),
              ...(container.history !== undefined && { history: container.history }),
              ...(container.docLocation !== undefined && {
                docLocation: container.docLocation,
              }),
              _docxHyperlinkIndex: targetInlineOccurrence(tr, container.occurrence),
            }),
        );
        break;
      }
      default: {
        const unreachable: never = container;
        return panic("Unhandled DOCX inline container during execution", {
          container: unreachable,
        });
      }
    }
  }
};

type ApplyTrackedFormattingOptions = ApplyDirectFormattingOptions &
  ExactFormattingMetadata & {
    readonly previousFormatting: Readonly<TextFormatting>;
    readonly revisionId: number;
  };

const applyExactTrackedFormatting = ({
  tr,
  from,
  to,
  formatting,
  previousFormatting,
  revisionId,
  author,
  date,
  styleResolver,
}: ApplyTrackedFormattingOptions): number => {
  if (from >= to) return revisionId;
  const propertyChangeType = tr.doc.type.schema.marks["runPropertyChange"];
  if (!propertyChangeType) {
    return panic("A preflighted DOCX comparison lost run-property-change support");
  }
  let nextRevisionId = revisionId;
  for (const representation of selectRunFormattingCarrierRepresentations({
    doc: tr.doc,
    from,
    to,
  })) {
    const context = paragraphRunStyleContextAt({
      doc: tr.doc,
      pos: representation.from,
      ...(styleResolver !== undefined && { styleResolver }),
    });
    const formattingMarks = reconcileRunFormattingMarks({
      authoredFormatting: formatting,
      context,
      node: representation.node,
      ...(styleResolver !== undefined && { styleResolver }),
    });
    const change: RunPropertyChange = {
      type: "runPropertyChange",
      info: { id: nextRevisionId++, author, date },
      ...(Object.keys(previousFormatting).length > 0 ? { previousFormatting } : {}),
    };
    const marks = propertyChangeType.create({ changes: [change] }).addToSet(formattingMarks);
    applyMarksToRunFormattingRepresentation({ tr, representation, marks });
  }
  return nextRevisionId;
};

export type AppliedDocxTextRange = {
  readonly transaction: Transaction;
  readonly revisionIds: readonly number[];
  readonly nextRevisionId: number;
};

export const applyPreflightedDocxTextRange = ({
  tr,
  preflight,
  revisionIdSeed,
  author,
  date,
  styleResolver,
  application,
}: ExactFormattingMetadata & {
  readonly tr: Transaction;
  readonly preflight: Extract<DocxTextRangePreflight, { readonly type: "ready" }>;
  readonly revisionIdSeed: number;
  readonly styleResolver: ReturnType<typeof getDocumentStyleResolver>;
  readonly application: "format" | "replace";
}): AppliedDocxTextRange => {
  const insertionType = tr.doc.type.schema.marks["insertion"];
  const deletionType = tr.doc.type.schema.marks["deletion"];
  if (!insertionType || !deletionType) {
    return panic("A preflighted DOCX comparison lost tracked-text schema support");
  }
  const revisionIds: number[] = [];
  let nextRevisionId = revisionIdSeed;
  const changesText = preflight.steps.some(({ type }) => type === "del" || type === "ins");
  if (application === "format" && changesText) {
    return panic("A formatting-only DOCX instruction contains a text mutation");
  }
  const deletionRevisionId = application === "replace" ? nextRevisionId++ : null;
  const insertionRevisionId = application === "replace" ? nextRevisionId++ : null;
  let usedDeletion = false;
  let usedInsertion = false;
  const targetInlineReuse = new Map(
    preflight.targetInlineReuse.map(({ occurrenceKey, mark }) => [occurrenceKey, mark]),
  );
  for (const step of preflight.steps.toReversed()) {
    if (step.type === "del") {
      if (deletionRevisionId === null) {
        return panic("A formatting-only DOCX instruction reached a deletion");
      }
      tr.addMark(
        step.from,
        step.to,
        deletionType.create({ revisionId: deletionRevisionId, author, date }),
      );
      usedDeletion = true;
      continue;
    }
    if (step.type === "ins") {
      if (insertionRevisionId === null) {
        return panic("A formatting-only DOCX instruction reached an insertion");
      }
      tr.insertText(step.text, step.at);
      tr.addMark(
        step.at,
        step.at + step.text.length,
        insertionType.create({ revisionId: insertionRevisionId, author, date }),
      );
      applyExactInlineOwnership({
        tr,
        from: step.at,
        to: step.at + step.text.length,
        containers: step.targetInlineContainers,
        targetInlineReuse,
      });
      applyExactDirectFormatting({
        tr,
        from: step.at,
        to: step.at + step.text.length,
        formatting: step.formatting,
        styleResolver,
      });
      usedInsertion = true;
      continue;
    }
    const before = nextRevisionId;
    nextRevisionId = applyExactTrackedFormatting({
      tr,
      from: step.from,
      to: step.to,
      formatting: step.fragment.targetFormatting,
      previousFormatting: step.fragment.sourceFormatting,
      revisionId: nextRevisionId,
      author,
      date,
      styleResolver,
    });
    for (let id = before; id < nextRevisionId; id++) revisionIds.push(id);
  }
  if (usedDeletion && deletionRevisionId !== null) revisionIds.unshift(deletionRevisionId);
  if (usedInsertion && insertionRevisionId !== null) {
    revisionIds.splice(usedDeletion ? 1 : 0, 0, insertionRevisionId);
  }
  return Object.freeze({
    transaction: tr,
    revisionIds: Object.freeze(revisionIds),
    nextRevisionId,
  });
};
