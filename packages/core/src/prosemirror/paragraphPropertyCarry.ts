/**
 * Handing one paragraph's properties to another, as a removed paragraph mark
 * requires.
 *
 * A paragraph's properties live on its mark, and a removed mark leaves the
 * paragraph after it: that is the paragraph left once the change is accepted.
 * An edit that should read as the FIRST paragraph (a merge, a Backspace at a
 * paragraph's start, a deletion across paragraphs) therefore gives the second
 * the first's properties, recorded as a property change under the same
 * revision, so accepting reads as the direct edit and rejecting restores the
 * second's own.
 */

import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import type { NumberingMap } from "../docx/numberingParser";
import { expectParagraphAttrs } from "./attrs";
import {
  hasSerializableParagraphPropertyChange,
  paragraphPropertiesSnapshot,
} from "./commands/propertyChangeScope";
import { resolveParagraphChangeAttrs } from "./commands/resolveParagraphProperties";
import { rebaseParagraphRuns } from "./rebaseParagraphRunFormatting";
import { paragraphRunStyleContext, type RunStyleResolver } from "./runStyleFormatting";
import type { ParagraphPropertyChangeAttrs } from "./schema/nodes";

type RecordedFormatting = NonNullable<ParagraphPropertyChangeAttrs["previousFormatting"]>;

type CarryParagraphPropertiesOptions = {
  tr: Transaction;
  /** The paragraph that takes the properties. */
  position: number;
  /** The paragraph whose properties it takes. */
  source: PMNode;
  styleResolver: RunStyleResolver | null;
  numbering: NumberingMap | null;
  /** Records the change as a tracked property change; applied directly when omitted. */
  revision?: ParagraphPropertyChangeAttrs["info"];
  /** Properties the paragraph keeps its own values of: ones its own batch set. */
  keep?: ReadonlySet<string>;
};

/**
 * Give a paragraph another one's properties: the pPr a property change
 * covers. Its identity, mark, mark run properties and section stay its own,
 * and its runs are re-read in the new style. `tracked` says whether a property
 * change now records what it had: a paragraph holds one, and a pending one
 * already records it.
 */
export const carryParagraphProperties = ({
  tr,
  position,
  source,
  styleResolver,
  numbering,
  revision,
  keep,
}: CarryParagraphPropertiesOptions): { tr: Transaction; changed: boolean; tracked: boolean } => {
  const target = tr.doc.nodeAt(position);
  if (!target || target.type !== source.type) {
    // No paragraph follows to take them (a bookmark boundary or a table):
    // the resolver keeps the mark there, and the first paragraph stays.
    return { tr, changed: false, tracked: false };
  }
  const previousFormatting = paragraphPropertiesSnapshot(target);
  const own = previousFormatting as Record<string, unknown>;
  const formatting = Object.fromEntries([
    ...Object.entries(paragraphPropertiesSnapshot(source)).filter(([key]) => !keep?.has(key)),
    ...Object.entries(own).filter(([key]) => keep?.has(key)),
  ]);
  if (JSON.stringify(previousFormatting) === JSON.stringify(formatting)) {
    return { tr, changed: false, tracked: false };
  }
  const existing = expectParagraphAttrs(target)._propertyChanges;
  const tracked = revision !== undefined && !hasSerializableParagraphPropertyChange(existing);
  // Restoring a property change sets exactly the in-scope properties and
  // leaves the rest, which is this hand-over read the other way round.
  const carried = resolveParagraphChangeAttrs({
    node: target.type.create(
      {
        ...target.attrs,
        _propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: { id: -1, author: "", date: "1970-01-01T00:00:00Z" },
            previousFormatting: formatting as RecordedFormatting,
          } satisfies ParagraphPropertyChangeAttrs,
        ],
      },
      target.content,
      target.marks,
    ),
    mode: "reject",
    boundaryCovered: true,
    revisionSet: null,
    styleResolver,
    numbering,
  });
  if (!carried) {
    return panic("Paragraph properties did not pass to the paragraph", { position });
  }
  const changes = [
    ...(Array.isArray(existing) ? existing : []),
    ...(tracked && revision
      ? [
          {
            type: "paragraphPropertyChange",
            info: revision,
            previousFormatting,
          } satisfies ParagraphPropertyChangeAttrs,
        ]
      : []),
  ];
  tr.setNodeMarkup(position, undefined, {
    ...carried,
    ...carriedNumberingProvenance(source, target, keep),
    // The hand-over covers pPr only: the section and its pending change stay.
    _sectionProperties: target.attrs["_sectionProperties"],
    _propertyChanges: changes.length > 0 ? changes : null,
  });
  if (styleResolver) {
    rebaseParagraphRuns({
      previousContext: paragraphRunStyleContext(target, styleResolver),
      paragraphPosition: position,
      styleResolver,
      tr,
    });
  }
  return { tr, changed: true, tracked };
};

/**
 * Whether the carried numbering is the paragraph's own or its style's: the
 * record holds only stated numbering, and a style's numbering
 * stays the style's rather than becoming the paragraph's own `w:numPr`.
 */
const carriedNumberingProvenance = (
  source: PMNode,
  target: PMNode,
  keep: ReadonlySet<string> | undefined,
): Record<string, unknown> => ({
  numPrFromStyle: expectParagraphAttrs(keep?.has("numPr") ? target : source).numPrFromStyle,
});

type ParagraphLeftAfterOptions = {
  doc: PMNode;
  paragraphPos: number;
  /** Breaks a batch's deferred final-paragraph deletion will remove. */
  removedBreakPositions?: ReadonlySet<number>;
};

/**
 * The paragraph left once the break of the paragraph at `paragraphPos` goes:
 * the next one, or, when its own break is pending deletion too, the first
 * paragraph after it whose break stays. Null when no paragraph follows.
 */
export const paragraphLeftAfter = ({
  doc,
  paragraphPos,
  removedBreakPositions,
}: ParagraphLeftAfterOptions): number | null => {
  let position = paragraphPos;
  const start = doc.nodeAt(position);
  if (!start) return null;
  let node: PMNode = start;
  for (;;) {
    const nextPos: number = position + node.nodeSize;
    const next: PMNode | null = doc.resolve(nextPos).nodeAfter;
    if (!next || next.type !== node.type) return null;
    const mark = expectParagraphAttrs(next).pPrMark;
    const nextGoes =
      removedBreakPositions?.has(nextPos) ||
      (mark != null && (mark.kind === "del" || mark.kind === "moveFrom"));
    const afterNext = doc.resolve(nextPos + next.nodeSize).nodeAfter;
    if (!nextGoes || afterNext?.type !== node.type) return nextPos;
    position = nextPos;
    node = next;
  }
};

const recordedFormatting = paragraphPropertiesSnapshot;

/**
 * A paragraph now holds the break another one had (a split, or a paste into
 * it, leaves that break on the last part) but reads with other properties:
 * record the other's as its property change. Rejecting the inserted breaks
 * leaves this paragraph, which then reads as the one that was there. A
 * pending property change on the other one travels as it is. Returns whether
 * a change under `revision` was written.
 */
export const recordReplacedParagraphProperties = ({
  tr,
  position,
  replaced,
  revision,
}: {
  tr: Transaction;
  position: number;
  replaced: PMNode;
  revision: ParagraphPropertyChangeAttrs["info"];
}): boolean => {
  const target = tr.doc.nodeAt(position);
  if (!target || target.type !== replaced.type) return false;
  if (hasSerializableParagraphPropertyChange(expectParagraphAttrs(target)._propertyChanges)) {
    return false;
  }
  const pending = expectParagraphAttrs(replaced)._propertyChanges;
  if (hasSerializableParagraphPropertyChange(pending)) {
    tr.setNodeAttribute(position, "_propertyChanges", pending);
    return false;
  }
  const previousFormatting = recordedFormatting(replaced);
  if (JSON.stringify(previousFormatting) === JSON.stringify(recordedFormatting(target))) {
    return false;
  }
  tr.setNodeAttribute(position, "_propertyChanges", [
    {
      type: "paragraphPropertyChange",
      info: revision,
      previousFormatting: previousFormatting as RecordedFormatting,
    } satisfies ParagraphPropertyChangeAttrs,
  ]);
  return true;
};

/**
 * The properties a batch's own property change set on a paragraph: the ones
 * whose value differs from what that change records as before.
 */
export const propertiesSetInBatch = (
  node: PMNode,
  batchRevisionIds: ReadonlySet<number>,
): ReadonlySet<string> => {
  const changes = expectParagraphAttrs(node)._propertyChanges;
  const change = Array.isArray(changes)
    ? changes.find(({ info }) => batchRevisionIds.has(info.id))
    : undefined;
  if (change === undefined) return new Set();
  const before = (change.previousFormatting ?? {}) as Record<string, unknown>;
  const now = paragraphPropertiesSnapshot(node) as Record<string, unknown>;
  return new Set(
    [...new Set([...Object.keys(before), ...Object.keys(now)])].filter(
      (key) => JSON.stringify(before[key]) !== JSON.stringify(now[key]),
    ),
  );
};
