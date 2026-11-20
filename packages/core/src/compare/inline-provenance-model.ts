/**
 * Write a comparison's inline provenance plan (`planInlineProvenance`) as
 * document operations on the reviewed story's model.
 *
 * The plan is read off the story's projection, so each change names editor
 * positions; the projection's own addresses turn them into text positions.
 * A retained stretch records its new formatting as a tracked property change,
 * the comparison's own insertion takes it outright, and a stretch whose link
 * changes is struck and written again after itself with the revised link.
 * Nothing here edits a document: the batch applies as one commit, in order.
 */

import {
  applyDocumentOp,
  type DeleteRangeOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertContentOp,
  type NewIds,
  revisionIdDemand,
  type RevisionStamp,
  type RunPropsPatch,
  type SetRunPropsOp,
  type TextPosition,
} from "@stll/docx-core/ops";

import type { FolioRevisionStamp } from "../ai-edits/apply";
import type { Hyperlink, Run, TextFormatting } from "../types/document";
import { readModel, type ReadonlyDocument } from "../types/documentModel";
import { canonicalJson } from "../utils/canonicalJson";
import {
  hyperlinkLinkOf,
  type PlannedInlineProvenance,
} from "./inline-provenance";

type InlineProvenanceOperationsOptions = {
  document: ReadonlyDocument;
  /** The plan, in document order. */
  changes: readonly PlannedInlineProvenance[];
  /** The text position an editor position of the story's projection names. */
  addressAt: (position: number) => TextPosition | null;
  revisionStamp: FolioRevisionStamp;
  author: string;
};

type InlineProvenanceOperationsResult =
  | {
      status: "matched";
      /** Applies as one batch, in order, to `document`. */
      operations: readonly DocumentOp[];
      nextRevisionId: number;
      changedTargetBlockIds: readonly string[];
    }
  | { status: "unalignable"; detail: string };

/** The keys whose values differ, each set to its revised value or removed. */
const formattingPatch = (
  previous: TextFormatting,
  next: TextFormatting,
): RunPropsPatch => {
  const before = new Map(Object.entries(previous));
  const after = new Map(Object.entries(next));
  const patch: RunPropsPatch = {};
  for (const key of new Set([...before.keys(), ...after.keys()])) {
    if (canonicalJson(before.get(key)) === canonicalJson(after.get(key)))
      continue;
    Reflect.set(patch, key, after.get(key) ?? null);
  }
  return patch;
};

/** A position between characters: zero-width children at it stay outside a range. */
const characterGap = ({
  story,
  blockId,
  offset,
}: TextPosition): TextPosition => ({ story, blockId, offset });

/** The link a stretch is written again inside, as the model holds a new one. */
const linkOf = (
  link: { href: string; tooltip?: string },
  run: Run,
): Hyperlink => ({
  type: "hyperlink",
  ...(link.href.startsWith("#")
    ? { anchor: link.href.slice(1) }
    : { href: link.href }),
  ...(link.tooltip === undefined ? {} : { tooltip: link.tooltip }),
  children: [run],
});

/** The plan as one batch of document operations, or why one cannot be written. */
export const inlineProvenanceOperations = ({
  document,
  changes,
  addressAt,
  revisionStamp,
  author,
}: InlineProvenanceOperationsOptions): InlineProvenanceOperationsResult => {
  const operations: DocumentOp[] = [];
  let current = readModel(document);
  const changedTargetBlockIds = new Set<string>();
  let nextRevisionId = revisionStamp.idSeed;
  let refusal: string | undefined;
  const stamp = (): RevisionStamp => ({
    id: nextRevisionId++,
    author,
    date: revisionStamp.date,
  });
  /** Apply one operation to the copy and keep it, or note why it cannot apply. */
  const emit = (op: DocumentOp): boolean => {
    if (refusal !== undefined) return false;
    const applied = applyDocumentOp(current, op);
    if (applied.isErr()) {
      refusal = applied.error.message;
      return false;
    }
    current = applied.value.document;
    operations.push(op);
    return true;
  };
  /** Emit an edit with the ids the records it cuts in two take, after its stamp's. */
  const emitWithFreshIds = (
    op: SetRunPropsOp | InsertContentOp | DeleteRangeOp,
  ): boolean => {
    const demand = revisionIdDemand(current, op);
    if (demand.isErr()) {
      refusal ??= demand.error.message;
      return false;
    }
    if (demand.value === 0) return emit(op);
    const newIds: NewIds = {
      revision: Array.from({ length: demand.value }, () => nextRevisionId++),
    };
    return emit({ ...op, newIds });
  };

  // Last first: a link written again after its stretch moves only what
  // follows it, which is already written.
  for (const change of changes.toReversed()) {
    const fromAddress = addressAt(change.live.from);
    const toAddress = addressAt(change.live.to);
    if (fromAddress === null || toAddress === null)
      return {
        status: "unalignable",
        detail: "a formatted stretch has no text position in the story",
      };
    const from = characterGap(fromAddress);
    const to = characterGap(toAddress);
    let written: boolean;
    if (change.hyperlinkChange === "replace") {
      const run: Run = {
        type: "run",
        ...(Object.keys(change.formatting).length > 0
          ? { formatting: change.formatting }
          : {}),
        content: [{ type: "text", text: change.live.node.textContent }],
      };
      const link = hyperlinkLinkOf(change.target.node);
      const content = link === undefined ? run : linkOf(link, run);
      // Retained text is struck and written again after itself; the
      // comparison's own insertion is replaced outright.
      written =
        emitWithFreshIds({
          type: DOCUMENT_OP_TYPES.DELETE_RANGE,
          from,
          to,
          ...(change.inserted ? {} : { revision: stamp() }),
        }) &&
        emitWithFreshIds({
          type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
          at: change.inserted ? from : to,
          slice: { content: [content], openStart: 0, openEnd: 0 },
          revision: stamp(),
        });
    } else {
      const patch = formattingPatch(
        change.previousFormatting,
        change.formatting,
      );
      written =
        Object.keys(patch).length === 0 ||
        emitWithFreshIds({
          type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
          from,
          to,
          patch,
          ...(change.inserted ? {} : { revision: stamp() }),
        });
    }
    if (!written)
      return {
        status: "unalignable",
        detail: `run formatting cannot be written as document operations: ${refusal ?? "an operation was refused"}`,
      };
    changedTargetBlockIds.add(change.targetBlockId);
  }
  return {
    status: "matched",
    operations,
    nextRevisionId,
    changedTargetBlockIds: [...changedTargetBlockIds],
  };
};
