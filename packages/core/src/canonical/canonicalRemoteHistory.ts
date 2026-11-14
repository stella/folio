import { panic } from "better-result";

import {
  applyBatch,
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentBatch,
  type DocumentOp,
  findStoryBody,
  type OpStory,
  paragraphLength,
  type SequencedBatch,
  storyParagraphs,
  type TextPosition,
  transformBatch,
  transformPosition,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";
import type { CanonicalSelection } from "./canonicalSession";

/**
 * A remote edit at a position's own offset lands after it: text another actor
 * types where the local caret stands appears after the caret, so typing goes
 * on where it was.
 */
const REMOTE_AFFINITY = "before";

/** Operations whose transform reads facts captured from the document. */
const EFFECT_OPERATION_TYPES: ReadonlySet<DocumentOp["type"]> = new Set([
  DOCUMENT_OP_TYPES.SPLIT_BLOCK,
  DOCUMENT_OP_TYPES.JOIN_BLOCKS,
  DOCUMENT_OP_TYPES.DELETE_BLOCKS,
  DOCUMENT_OP_TYPES.RESOLVE_REVISION,
]);
const readsEffects = (ops: readonly DocumentOp[]): boolean =>
  ops.some((op) => EFFECT_OPERATION_TYPES.has(op.type));

/**
 * A selection moved across `over` by paragraph id and offset. A cell or
 * inline-node selection becomes a text selection over the same ends, which
 * the remote edit may have reshaped. Undefined when an end cannot be mapped.
 * The session moves its own selection this way and a peer's presence moves
 * the same way, so both land where the other expects.
 */
export const remapSelection = (
  selection: CanonicalSelection,
  over: SequencedBatch,
): CanonicalSelection | undefined => {
  const anchor = transformPosition(selection.anchor, [over], REMOTE_AFFINITY);
  const head = transformPosition(selection.head, [over], REMOTE_AFFINITY);
  if (anchor === undefined || head === undefined) return undefined;
  return { type: selection.type === "all" ? "all" : "text", anchor, head };
};

const storyParagraphList = (document: Document, story: OpStory) => {
  const body = findStoryBody(document, story);
  return body === undefined
    ? []
    : storyParagraphs(body).map(({ paragraph }) => paragraph);
};

/**
 * Where a position the remote edit could not map goes: its paragraph clamped
 * to the paragraph's new length, else the start of the nearest paragraph
 * that survived, following ones first, else the start of the story.
 */
const nearestPosition = ({
  position,
  previous,
  document,
  story,
}: {
  position: TextPosition | undefined;
  previous: Document;
  document: Document;
  story: OpStory;
}): TextPosition => {
  const target = position?.story ?? story;
  const paragraphs = storyParagraphList(document, target);
  const surviving = new Map(
    paragraphs.flatMap((paragraph) =>
      paragraph.paraId === undefined
        ? []
        : [[paragraph.paraId, paragraph] as const],
    ),
  );
  if (position !== undefined) {
    const same = surviving.get(position.blockId);
    if (same !== undefined)
      return {
        story: target,
        blockId: position.blockId,
        offset: Math.min(position.offset, paragraphLength(same)),
      };
    const order = storyParagraphList(previous, target).map(
      ({ paraId }) => paraId,
    );
    const index = order.indexOf(position.blockId);
    if (index >= 0) {
      const candidates = [
        ...order.slice(index + 1),
        ...order.slice(0, index).toReversed(),
      ];
      const nearest = candidates.find(
        (blockId) => blockId !== undefined && surviving.has(blockId),
      );
      if (nearest !== undefined)
        return { story: target, blockId: nearest, offset: 0 };
    }
  }
  const first = paragraphs.find(({ paraId }) => paraId !== undefined)?.paraId;
  if (first === undefined)
    return panic("A story keeps an identified paragraph.");
  return { story: target, blockId: first, offset: 0 };
};

/** The local selection after a remote edit, moved by paragraph id and offset. */
export const remoteSelection = ({
  selection,
  over,
  previous,
  document,
  story,
}: {
  selection: CanonicalSelection | null;
  over: SequencedBatch;
  previous: Document;
  document: Document;
  story: OpStory;
}): CanonicalSelection => {
  const mapped =
    selection === null ? undefined : remapSelection(selection, over);
  if (mapped !== undefined) return mapped;
  const caret = nearestPosition({
    position: selection?.head,
    previous,
    document,
    story,
  });
  return {
    type: selection?.type === "all" ? "all" : "text",
    anchor: caret,
    head: caret,
  };
};

/**
 * One step of local history: `ops` apply to the document the step above it
 * left (the top step: the current document), `before` is the selection in
 * that document and `after` the one the step leaves.
 */
export type HistoryStep = {
  ops: readonly DocumentOp[];
  before: CanonicalSelection;
  after: CanonicalSelection;
};

const HISTORY_BATCH = {
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId: "local-history",
  actor: "local-history",
} as const;

/**
 * Rebase a history stack, top step first, over a remote edit made to the
 * document the top step applies to. Each step's ops are transformed over the
 * remote edit, then the remote edit over the step, so the next step down sees
 * it in its own coordinates. The stack stops at the first step that does not
 * rebase: it and every step beneath it are left out of the result, since
 * they were recorded on top of what it restores.
 */
export const rebaseHistorySteps = ({
  document,
  over,
  steps,
  fallback,
}: {
  document: Document;
  over: SequencedBatch;
  steps: readonly HistoryStep[];
  /** The selection a step takes when the remote edit removed its own. */
  fallback: CanonicalSelection;
}): HistoryStep[] => {
  let remote = over;
  // The document the next step applies to, reached lazily: only operations
  // whose transform reads captured facts need it.
  let cursor = document;
  let unapplied: (readonly DocumentOp[])[] = [];
  const settled = (): Document | undefined => {
    for (const ops of unapplied) {
      const applied = applyBatch(cursor, ops);
      if (applied.isErr()) return undefined;
      cursor = applied.value.document;
    }
    unapplied = [];
    return cursor;
  };
  const rebased: HistoryStep[] = [];
  for (const step of steps) {
    const batch: DocumentBatch = {
      ...HISTORY_BATCH,
      baseRev: remote.baseRev,
      ops: step.ops,
    };
    const adjusted = transformBatch(batch, [remote]);
    if (adjusted.isErr()) break;
    const before = remapSelection(step.before, remote) ?? fallback;
    let effects: SequencedBatch["effects"];
    if (readsEffects(step.ops) || readsEffects(remote.ops)) {
      const base = settled();
      if (base === undefined) break;
      const applied = applyBatch(base, step.ops);
      if (applied.isErr()) break;
      cursor = applied.value.document;
      effects = applied.value.effects;
    } else unapplied.push(step.ops);
    const propagated = transformBatch(
      remote,
      [
        {
          ...batch,
          revision: remote.revision,
          ...(effects === undefined ? {} : { effects }),
        },
      ],
      { order: "before" },
    );
    if (propagated.isErr()) {
      rebased.push({ ops: adjusted.value.ops, before, after: fallback });
      break;
    }
    let remoteEffects: SequencedBatch["effects"];
    if (readsEffects(propagated.value.ops)) {
      const base = settled();
      const captured =
        base === undefined ? undefined : applyBatch(base, propagated.value.ops);
      if (captured === undefined || captured.isErr()) {
        rebased.push({ ops: adjusted.value.ops, before, after: fallback });
        break;
      }
      remoteEffects = captured.value.effects;
    }
    remote = {
      ...propagated.value,
      revision: remote.revision,
      ...(remoteEffects === undefined ? {} : { effects: remoteEffects }),
    };
    rebased.push({
      ops: adjusted.value.ops,
      before,
      after: remapSelection(step.after, remote) ?? fallback,
    });
  }
  return rebased;
};
