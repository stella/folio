import { panic, Result } from "better-result";

import {
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  SPLIT_HALVES,
  type DocumentOp,
  type TextPosition,
} from "../types";
import {
  BATCH_REJECTION_REASONS,
  BatchRejection,
  type DocumentBatch,
  type SequencedBatch,
  type SequencedOpEffect,
} from "./envelope";
import { sameStory } from "./address";

type TransformOptions = { order?: "before" | "after" };

const operationStory = (op: DocumentOp) => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
      return op.at.story;
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return op.from.story;
    case DOCUMENT_OP_TYPES.DELETE_BLOCKS:
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
    case DOCUMENT_OP_TYPES.INSERT_TABLE:
    case DOCUMENT_OP_TYPES.DELETE_TABLE:
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
    case DOCUMENT_OP_TYPES.INSERT_ROW:
    case DOCUMENT_OP_TYPES.DELETE_ROW:
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
      return op.story;
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
      return undefined;
    default: {
      const exhaustive: never = op;
      void exhaustive;
      return panic("Unknown document operation kind.");
    }
  }
};

type ParagraphAddress = Pick<TextPosition, "story" | "blockId">;
const sameParagraph = (left: ParagraphAddress, right: ParagraphAddress): boolean =>
  sameStory(left.story, right.story) && left.blockId === right.blockId;

const targets = (op: DocumentOp): readonly string[] | undefined => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
      return [op.at.blockId];
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return [op.from.blockId, op.to.blockId];
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
      return [op.blockId];
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
      return [op.blockId, op.nextBlockId];
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
      return [op.at.blockId, ...op.blocks.flatMap((p) => (p.paraId ? [p.paraId] : []))];
    default:
      return undefined;
  }
};

const independent = (a: DocumentOp, b: DocumentOp): boolean => {
  const left = targets(a);
  const right = targets(b);
  return left !== undefined && right !== undefined && !left.some((id) => right.includes(id));
};

const refusal = (op: DocumentOp, over: DocumentOp, message: string) =>
  Result.err(
    new BatchRejection({
      reason: BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
      message,
      opType: op.type,
      overType: over.type,
    }),
  );

type PairOptions = {
  op: DocumentOp;
  over: DocumentOp;
  effect: SequencedOpEffect | undefined;
  order: "before" | "after";
};
const transformOp = ({
  op,
  over,
  effect,
  order,
}: PairOptions): Result<readonly DocumentOp[], BatchRejection> => {
  const localStory = operationStory(op);
  const remoteStory = operationStory(over);
  if (localStory === undefined || remoteStory === undefined)
    return refusal(op, over, "Story lifecycle changes have no supported sequencing transform.");
  if (!sameStory(localStory, remoteStory)) return Result.ok([op]);
  if (independent(op, over)) return Result.ok([op]);
  if (effect?.type === "touchedBlocks" && op.type !== DOCUMENT_OP_TYPES.RESOLVE_REVISION) {
    const addressed = targets(op);
    if (addressed && addressed.every((id) => !effect.blockIds.includes(id))) return Result.ok([op]);
    return refusal(op, over, "The operation targets a structurally changed block.");
  }
  if (
    op.type === DOCUMENT_OP_TYPES.RESOLVE_REVISION &&
    over.type === DOCUMENT_OP_TYPES.RESOLVE_REVISION
  ) {
    if (
      op.decision !== over.decision &&
      op.revisionIds.some((id) => over.revisionIds.includes(id))
    ) {
      return refusal(op, over, "Conflicting revision decisions.");
    }
    return Result.ok([op]);
  }
  if (
    op.type === DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS &&
    over.type === DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS
  ) {
    if (
      "numPr" in op.patch ||
      "numPr" in over.patch ||
      op.expected ||
      over.expected ||
      op.revision ||
      over.revision
    ) {
      return refusal(
        op,
        over,
        "Numbering, tracked properties and property preconditions require an exclusive edit.",
      );
    }
    if (order === "before") {
      const patch = { ...op.patch };
      for (const key of Object.keys(over.patch)) Reflect.deleteProperty(patch, key);
      return Result.ok([{ ...op, patch }]);
    }
    return Result.ok([op]);
  }
  if (
    (op.type === DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS &&
      (over.type === DOCUMENT_OP_TYPES.INSERT_TEXT ||
        over.type === DOCUMENT_OP_TYPES.DELETE_RANGE ||
        over.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS)) ||
    (over.type === DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS &&
      (op.type === DOCUMENT_OP_TYPES.INSERT_TEXT ||
        op.type === DOCUMENT_OP_TYPES.DELETE_RANGE ||
        op.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS))
  )
    return Result.ok([op]);
  if (
    op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS &&
    (over.type === DOCUMENT_OP_TYPES.INSERT_TEXT ||
      over.type === DOCUMENT_OP_TYPES.DELETE_RANGE ||
      over.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS)
  )
    return Result.ok([op]);
  if (
    op.type === DOCUMENT_OP_TYPES.INSERT_BLOCKS &&
    over.type === DOCUMENT_OP_TYPES.INSERT_BLOCKS
  ) {
    if (op.at.type !== over.at.type || op.at.blockId !== over.at.blockId)
      return refusal(op, over, "Overlapping block insertion identities.");
    if ((order === "before") === (op.at.type === "after")) return Result.ok([op]);
    const anchor = op.at.type === "after" ? over.blocks.at(-1) : over.blocks.at(0);
    if (!anchor?.paraId) return refusal(op, over, "Inserted paragraph has no identity.");
    return Result.ok([{ ...op, at: { type: op.at.type, blockId: anchor.paraId } }]);
  }
  if (
    op.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS &&
    over.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS
  ) {
    if (op.expected || over.expected || op.revision || over.revision)
      return refusal(op, over, "Tracked or conditional run properties need an unchanged range.");
    if (
      Object.keys(op.patch).some((key) => key in over.patch) &&
      Math.max(op.from.offset, over.from.offset) < Math.min(op.to.offset, over.to.offset) &&
      (op.from.offset !== over.from.offset || op.to.offset !== over.to.offset)
    )
      return refusal(
        op,
        over,
        "Partially overlapping competing keys require captured run boundaries.",
      );
    if (order === "before" && Object.keys(op.patch).some((key) => key in over.patch)) {
      if (Math.max(op.from.offset, over.from.offset) >= Math.min(op.to.offset, over.to.offset))
        return Result.ok([op]);
      const patch = { ...op.patch };
      for (const key of Object.keys(over.patch)) Reflect.deleteProperty(patch, key);
      return Result.ok(Object.keys(patch).length > 0 ? [{ ...op, patch }] : []);
    }
    return Result.ok([op]);
  }
  if (
    op.type !== DOCUMENT_OP_TYPES.INSERT_TEXT &&
    op.type !== DOCUMENT_OP_TYPES.DELETE_RANGE &&
    op.type !== DOCUMENT_OP_TYPES.SET_RUN_PROPS &&
    op.type !== DOCUMENT_OP_TYPES.SPLIT_BLOCK
  ) {
    return refusal(op, over, "This operation pair has no position transform.");
  }
  if (
    (op.type === DOCUMENT_OP_TYPES.DELETE_RANGE && (op.expected || op.join !== undefined)) ||
    (op.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS &&
      (op.expected || op.joinStart !== undefined || op.joinEnd !== undefined))
  ) {
    let outside = false;
    if (over.type === DOCUMENT_OP_TYPES.INSERT_TEXT) {
      outside = over.at.offset < op.from.offset || over.at.offset > op.to.offset;
    }
    if (over.type === DOCUMENT_OP_TYPES.DELETE_RANGE) {
      outside = over.to.offset < op.from.offset || over.from.offset > op.to.offset;
    }
    if (!outside)
      return refusal(op, over, "Structural inverse preconditions require an unchanged range.");
  }
  if (
    op.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS &&
    over.type === DOCUMENT_OP_TYPES.INSERT_TEXT &&
    over.runProps !== INHERIT_RUN_PROPS &&
    over.at.offset >= op.from.offset &&
    over.at.offset <= op.to.offset
  )
    return refusal(op, over, "Explicit inserted formatting needs a range partition.");
  if (
    op.type === DOCUMENT_OP_TYPES.INSERT_TEXT &&
    over.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS &&
    op.runProps !== INHERIT_RUN_PROPS &&
    op.at.offset >= over.from.offset &&
    op.at.offset <= over.to.offset
  )
    return refusal(op, over, "Explicit inserted formatting needs a range partition.");
  if (
    over.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
    over.revision &&
    op.type === DOCUMENT_OP_TYPES.INSERT_TEXT &&
    op.at.offset > over.from.offset &&
    op.at.offset < over.to.offset
  )
    return refusal(op, over, "Insertion inside tracked-deleted content requires wrapper affinity.");
  if (
    op.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
    op.revision &&
    over.type === DOCUMENT_OP_TYPES.INSERT_TEXT &&
    over.at.offset > op.from.offset &&
    over.at.offset < op.to.offset
  )
    return refusal(op, over, "Tracking a concurrent insertion requires a revision partition.");
  if (
    over.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
    ((over.from.zeroWidthBefore ?? 0) !== 0 || (over.to.zeroWidthBefore ?? 0) !== 0)
  ) {
    return refusal(
      op,
      over,
      "Tracked deletion and zero-width boundaries require an unchanged range.",
    );
  }
  const map = (
    point: TextPosition,
    affinity: "before" | "after" = "after",
  ): TextPosition | undefined => {
    if (over.type === DOCUMENT_OP_TYPES.INSERT_TEXT) {
      if (!sameParagraph(point, over.at)) return point;
      const shift =
        point.offset > over.at.offset || (point.offset === over.at.offset && order === "after");
      return shift ? { ...point, offset: point.offset + over.text.length } : point;
    }
    if (over.type === DOCUMENT_OP_TYPES.DELETE_RANGE) {
      if (!sameParagraph(point, over.from)) return point;
      if (over.revision) return point;
      if (
        (point.zeroWidthBefore ?? 0) !== 0 &&
        point.offset >= over.from.offset &&
        point.offset <= over.to.offset
      )
        return undefined;
      return {
        ...point,
        offset:
          point.offset <= over.from.offset
            ? point.offset
            : Math.max(over.from.offset, point.offset - (over.to.offset - over.from.offset)),
      };
    }
    if (over.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK) {
      if (!sameParagraph(point, over.at)) return point;
      const half =
        over.newHalf ??
        (effect?.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK ? effect.newHalf : undefined);
      if (!half || (point.zeroWidthBefore ?? 0) !== 0 || (over.at.zeroWidthBefore ?? 0) !== 0)
        return undefined;
      const second =
        point.offset > over.at.offset ||
        (point.offset === over.at.offset &&
          affinity === "after" &&
          (op.type !== DOCUMENT_OP_TYPES.SPLIT_BLOCK || order === "after"));
      return {
        ...point,
        blockId: second === (half === SPLIT_HALVES.SECOND) ? over.newBlockId : over.at.blockId,
        offset: second ? point.offset - over.at.offset : point.offset,
      };
    }
    if (over.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS) {
      if (
        !sameStory(point.story, over.story) ||
        (point.blockId !== over.blockId && point.blockId !== over.nextBlockId)
      )
        return point;
      if (over.revision) return point;
      const length =
        effect?.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS ? effect.firstLength : undefined;
      if (length === undefined) return undefined;
      const survivor = over.survivor === SPLIT_HALVES.FIRST ? over.blockId : over.nextBlockId;
      return {
        ...point,
        blockId: survivor,
        offset: point.offset + (point.blockId === over.nextBlockId ? length : 0),
      };
    }
    if (over.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS) return point;
    return undefined;
  };
  if (op.type === DOCUMENT_OP_TYPES.INSERT_TEXT || op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK) {
    if (op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK && over.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS)
      return refusal(op, over, "Splitting a joined boundary requires an exclusive edit.");
    if (
      op.type === DOCUMENT_OP_TYPES.INSERT_TEXT &&
      over.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS &&
      !over.revision &&
      (over.depth ?? 0) === 0 &&
      op.at.blockId === over.nextBlockId &&
      op.at.offset === 0
    )
      return refusal(op, over, "Insertion at a retained run boundary needs explicit affinity.");
    const at = map(op.at);
    return at
      ? Result.ok([{ ...op, at }])
      : refusal(op, over, "Position cannot be mapped without structural context.");
  }
  // Keep a concurrent insertion out of a direct deletion, even inside its range.
  if (
    op.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
    !op.revision &&
    over.type === DOCUMENT_OP_TYPES.INSERT_TEXT &&
    op.from.offset < over.at.offset &&
    over.at.offset < op.to.offset
  ) {
    if (op.newIds || over.revision)
      return refusal(op, over, "Splitting this deletion requires new revision identities.");
    return Result.ok([
      {
        ...op,
        from: { ...op.from, offset: over.at.offset + over.text.length },
        to: { ...op.to, offset: op.to.offset + over.text.length },
      },
      { ...op, to: { ...op.to, offset: over.at.offset } },
    ]);
  }
  const from = map(op.from);
  let to = map(op.to, "before");
  if (
    over.type === DOCUMENT_OP_TYPES.INSERT_TEXT &&
    op.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS &&
    op.to.offset === over.at.offset &&
    to
  )
    to = { ...to, offset: op.to.offset + over.text.length };
  const adjustedFrom =
    over.type === DOCUMENT_OP_TYPES.INSERT_TEXT &&
    ((op.type === DOCUMENT_OP_TYPES.DELETE_RANGE && !op.revision) ||
      op.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS) &&
    op.from.offset === over.at.offset &&
    from
      ? { ...from, offset: op.from.offset + over.text.length }
      : from;
  if (
    over.type === DOCUMENT_OP_TYPES.INSERT_TEXT &&
    op.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
    !op.revision &&
    op.to.offset === over.at.offset &&
    to
  )
    to = { ...to, offset: op.to.offset };
  if (!adjustedFrom || !to)
    return refusal(op, over, "Range cannot be mapped without structural context.");
  if (!sameParagraph(adjustedFrom, to)) {
    if (!sameStory(adjustedFrom.story, to.story))
      return refusal(op, over, "A range cannot cross stories.");
    if (
      over.type !== DOCUMENT_OP_TYPES.SPLIT_BLOCK ||
      op.newIds ||
      op.from.offset >= over.at.offset ||
      op.to.offset <= over.at.offset
    )
      return refusal(op, over, "Range crosses an unsupported paragraph boundary.");
    const first = { ...op, from: adjustedFrom, to: { ...adjustedFrom, offset: over.at.offset } };
    const second = { ...op, from: { ...to, offset: 0 }, to };
    return Result.ok(
      op.type === DOCUMENT_OP_TYPES.DELETE_RANGE ? [second, first] : [first, second],
    );
  }
  if (adjustedFrom.offset === to.offset && op.type === DOCUMENT_OP_TYPES.DELETE_RANGE)
    return Result.ok([]);
  return Result.ok([{ ...op, from: adjustedFrom, to }]);
};

/** Rebase an atomic batch over an ordered journal tail. */
export const transformBatch = (
  batch: DocumentBatch,
  over: readonly SequencedBatch[],
  options: TransformOptions = {},
): Result<DocumentBatch, BatchRejection> => {
  for (const candidate of [batch, ...over]) {
    const invalid = candidate.ops.find(
      (op) =>
        (op.type === DOCUMENT_OP_TYPES.DELETE_RANGE ||
          op.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS) &&
        !sameParagraph(op.from, op.to),
    );
    if (invalid === undefined) continue;
    return Result.err(
      new BatchRejection({
        reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
        message: "A range must start and end in the same paragraph.",
        opType: invalid.type,
      }),
    );
  }
  let ops = [...batch.ops];
  const order = options.order ?? "after";
  for (const sequenced of over) {
    for (const [index, other] of sequenced.ops.entries()) {
      const next: DocumentOp[] = [];
      let remote: readonly DocumentOp[] = [other];
      let remoteEffect = sequenced.effects?.at(index);
      for (const [opIndex, op] of ops.entries()) {
        let local: readonly DocumentOp[] = [op];
        for (const remoteOp of remote) {
          const mapped: DocumentOp[] = [];
          for (const localOp of local) {
            const transformed = transformOp({
              op: localOp,
              over: remoteOp,
              effect: remoteEffect,
              order,
            });
            if (transformed.isErr()) return transformed;
            mapped.push(...transformed.value);
          }
          local = mapped;
        }
        next.push(...local);
        if (opIndex === ops.length - 1) continue;
        if (remote.length > 1)
          return refusal(op, other, "An expanded foreign deletion requires a sequence partition.");
        // The next local operation is addressed after this original local
        // operation. Move the foreign operation into that coordinate space.
        const reciprocal: DocumentOp[] = [];
        for (const remoteOp of remote) {
          const localStory = operationStory(op);
          const mapped = transformOp({
            op: remoteOp,
            over: op,
            effect: undefined,
            order: order === "after" ? "before" : "after",
          });
          if (mapped.isErr()) return mapped;
          reciprocal.push(...mapped.value);
          if (
            remoteOp.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS &&
            !remoteOp.revision &&
            remoteEffect?.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS &&
            localStory !== undefined &&
            sameStory(localStory, remoteOp.story)
          ) {
            if (op.type === DOCUMENT_OP_TYPES.INSERT_TEXT && op.at.blockId === remoteOp.blockId) {
              remoteEffect = {
                type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
                firstLength: remoteEffect.firstLength + op.text.length,
              };
            }
            if (
              op.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
              !op.revision &&
              op.from.blockId === remoteOp.blockId &&
              op.to.blockId === remoteOp.blockId
            ) {
              remoteEffect = {
                type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
                firstLength: remoteEffect.firstLength - (op.to.offset - op.from.offset),
              };
            }
          }
        }
        remote = reciprocal;
      }
      ops = next;
    }
  }
  const last = over.at(-1);
  return Result.ok({ ...batch, baseRev: last?.revision ?? batch.baseRev, ops });
};
