/** Paragraph insertion, isolated from the inline and table operation implementations. */
import { Result } from "better-result";

import {
  type Document,
  finalizeParagraphFormatting,
  type Paragraph,
} from "../model/document";
import {
  markBeforeTableResolves,
  paragraphTableRunInBoundary,
} from "./blockDeletion";
import {
  blockListAt,
  blockListEndsInParagraph,
  hasContainerFinalParagraphMark,
  replaceParagraphs,
  storyBody,
  storyParagraphs,
} from "./blocks";
import { validateOpsDocument } from "./contract";
import type { DocumentEdit } from "./edits";
import { structurallyEqual } from "./equality";
import { freshenIdentities } from "./identity";
import { identityKeysIn, idKey, isParaId, packageIdentityKeys } from "./ids";
import { asMoveHalf } from "./moves";
import { paragraphLength } from "./offsets";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { ApplyOps } from "./resolve";
import { stampInfo, WRAP_KINDS, wrapTracked } from "./review";
import { locateBlockInsertion } from "./tableLocation";
import {
  DOCUMENT_OP_TYPES,
  type InsertBlocksOp,
  type RevisionStamp,
} from "./types";

const sameFormatting = (
  left: Paragraph["formatting"],
  right: Paragraph["formatting"],
): boolean =>
  structurallyEqual(
    finalizeParagraphFormatting(left) ?? {},
    finalizeParagraphFormatting(right) ?? {},
  );

type InsertedContentOptions = {
  paragraph: Paragraph;
  revision: RevisionStamp;
  move: InsertBlocksOp["move"];
};
const insertedContent = ({
  paragraph,
  revision,
  move,
}: InsertedContentOptions): Result<Paragraph["content"], DocumentOpRefusal> => {
  const wrapped = wrapTracked({
    items: paragraph.content,
    from: { offset: 0, zeroWidthBefore: 0 },
    to: {
      offset: paragraphLength(paragraph),
      zeroWidthBefore: Number.MAX_SAFE_INTEGER,
    },
    kind: WRAP_KINDS.INSERTION,
    stamp: revision,
  });
  if (wrapped.kind === "refused")
    return Result.err(
      new DocumentOpRefusal({
        opType: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        reason: wrapped.reason,
        message: "The inserted content cannot be tracked.",
      }),
    );
  const content =
    wrapped.kind === "wrapped" ? wrapped.content : paragraph.content;
  return Result.ok(
    move === true ? asMoveHalf(content, revision.id, "to") : content,
  );
};

type BlockInsertionOptions = {
  document: Document;
  op: InsertBlocksOp;
  applyOps: ApplyOps;
};

export const insertBlocks = ({
  document,
  op,
  applyOps,
}: BlockInsertionOptions): Result<DocumentEdit, DocumentOpRefusal> => {
  const refuse = (reason: DocumentOpRefusal["reason"], message: string) =>
    Result.err(new DocumentOpRefusal({ reason, message, opType: op.type }));
  const firstInserted = op.blocks.at(0);
  if (firstInserted === undefined) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
      "No paragraphs to insert.",
    );
  }
  if (op.at.type === "beforeTable" || op.at.type === "afterTable") {
    if (
      op.blocks.some(({ sectionProperties }) => sectionProperties !== undefined)
    )
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        "Inserting section breaks requires section operations.",
      );
    if (
      op.blocks.some(({ paraId }) => paraId === undefined || !isParaId(paraId))
    )
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
        "Every inserted paragraph needs a usable id.",
      );
    const located = locateBlockInsertion(document, op);
    if (located.isErr()) return located;
    let paragraphs = op.blocks.map((paragraph) => {
      if (finalizeParagraphFormatting(paragraph.formatting) !== undefined)
        return paragraph;
      const canonical = { ...paragraph };
      delete canonical.formatting;
      return canonical;
    });
    if (op.revision !== undefined) {
      const list = blockListAt(
        storyBody(document, op.story).content,
        located.value.list,
      );
      // Independent review can carry an earlier first-source mark across
      // intervening blocks. Insertion cannot record that conditional ownership.
      for (let index = 0; index < located.value.index; index += 1) {
        const previous = list[index];
        if (
          previous?.type === "paragraph" &&
          previous.pPrMark?.survivor === "first"
        )
          return refuse(
            DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            "Tracked table-side insertion cannot cross an earlier pending first-paragraph owner.",
          );
      }
      const following = list.at(located.value.index);
      if (!list.some((block) => block.type === "paragraph"))
        return refuse(
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          "Tracked table-side insertion needs a surviving paragraph for its review inverse.",
        );
      if (following?.type === "table") {
        const boundary = paragraphTableRunInBoundary(firstInserted, following);
        if (boundary.type === "blocked" && boundary.reason !== "review")
          return refuse(
            DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            `Tracked table-side insertion cannot resolve a ${boundary.reason} boundary.`,
          );
        if (boundary.type === "receiver" && !blockListEndsInParagraph(list))
          return refuse(
            DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            "Tracked table-side insertion needs a paragraph terminator after its table receiver.",
          );
      } else if (following?.type !== "paragraph")
        return refuse(
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          "Tracked table-side insertion needs a following paragraph or a proven table boundary.",
        );
      if (identityKeysIn(op.blocks).length > 0)
        return refuse(
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          "Tracked insertion of identified review or content-control records is unsupported.",
        );
      const tracked: Paragraph[] = [];
      for (const paragraph of paragraphs) {
        const content = insertedContent({
          paragraph,
          revision: op.revision,
          move: op.move,
        });
        if (content.isErr()) return content;
        tracked.push({
          ...paragraph,
          content: content.value,
          pPrMark: {
            kind: op.move === true ? "moveTo" : "ins",
            info: stampInfo(op.revision),
          },
        });
      }
      const fresh = freshenIdentities({
        before: [],
        after: tracked,
        newIds: op.newIds ?? {},
        usedElsewhere: () => new Set(packageIdentityKeys(document.package)),
      });
      switch (fresh.kind) {
        case "fresh":
          paragraphs = fresh.paragraphs;
          break;
        case "needsIds":
          return refuse(
            DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
            "Inserted paragraph review records need more revision ids.",
          );
        case "invalidId":
          return refuse(
            DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
            "An inserted review record has an invalid revision id.",
          );
        default: {
          const unreachable: never = fresh;
          return unreachable;
        }
      }
    }
    const inserted = replaceParagraphs({
      document,
      story: op.story,
      at: located.value,
      count: 0,
      replacement: paragraphs,
    });
    if (inserted.isErr())
      return refuse(inserted.error.reason, inserted.error.message);
    const insertedBody = storyBody(inserted.value, op.story);
    if (
      paragraphs.some((paragraph, index) =>
        hasContainerFinalParagraphMark(insertedBody, {
          list: located.value.list,
          index: located.value.index + index,
          paragraph,
        }),
      )
    )
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
        "Insertion would leave a tracked mark on a container-final paragraph.",
      );
    const valid = validateOpsDocument(inserted.value);
    if (valid.isErr()) return refuse(valid.error.reason, valid.error.message);
    return Result.ok({
      document: inserted.value,
      inverse: [
        {
          type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
          story: op.story,
          expected: paragraphs,
          blocks: [],
        },
      ],
      touched: {
        modified: [],
        inserted: paragraphs.flatMap(({ paraId }) =>
          paraId === undefined ? [] : [paraId],
        ),
        removed: [],
      },
    });
  }
  if (op.at.type !== "before" && op.at.type !== "after") {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "Insert before or after a paragraph or table.",
    );
  }
  const body = storyBody(document, op.story);
  const location = storyParagraphs(body).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(op.at.blockId),
  );
  if (location === undefined) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "The insertion anchor does not exist.",
    );
  }
  if (
    op.blocks.some(({ sectionProperties }) => sectionProperties !== undefined)
  ) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "Inserting section breaks requires section operations.",
    );
  }
  // Use the following paragraph as the unchanged boundary whenever possible.
  const list = blockListAt(body.content, location.list);
  const following = list[location.index + 1];
  const before = op.at.type === "before" || following?.type === "paragraph";
  const anchor =
    op.at.type === "after" && following?.type === "paragraph"
      ? following
      : location.paragraph;
  // With no paragraph after the anchor, a tracked insertion is recorded as
  // a reader reads it: the anchor's mark is the inserted one, and the last
  // new paragraph holds the mark the anchor had, with the anchor's paragraph
  // properties as its prior ones. Rejecting removes the new text and the
  // anchor's mark, which leaves the anchor as it was; accepting leaves both.
  // Before a table the new paragraph's own mark can be inserted only when
  // the shared resolution owner admits it. Otherwise keep the source mark
  // on the new last paragraph, so rejecting joins back into the anchor.
  // At the end of the list, before a range boundary between blocks that no
  // resolution may carry content across, or before any other block, such as
  // a content control, there is no mark after the new paragraph's for it to
  // resolve against, so the anchor's mark is the inserted one.
  const takesOverMark =
    op.revision !== undefined &&
    !before &&
    (following?.type !== "table" ||
      !markBeforeTableResolves({
        blocks: list,
        boundary: paragraphTableRunInBoundary(firstInserted, following),
      }));
  if (
    takesOverMark &&
    (location.paragraph.pPrMark !== undefined ||
      location.paragraph.reviewCarrier !== undefined ||
      (location.paragraph.propertyChanges?.length ?? 0) > 0)
  ) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "A tracked insertion after a paragraph no paragraph follows needs the anchor's mark free of review records.",
    );
  }
  const replacement = (
    blocks: readonly Paragraph[],
    replacedAnchor: Paragraph = anchor,
  ) => ({
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: op.story,
    expected: [anchor],
    blocks: before ? [...blocks, replacedAnchor] : [replacedAnchor, ...blocks],
  });
  const canonicalBlocks = op.blocks.map((paragraph) => {
    if (finalizeParagraphFormatting(paragraph.formatting) !== undefined)
      return paragraph;
    const canonical = { ...paragraph };
    delete canonical.formatting;
    return canonical;
  });
  // A direct paragraph after this pending break would be joined into its
  // anchor when the mark is rejected. Moving the break instead needs the
  // former mark's formatting snapshot and revision ids, which this direct
  // insertion does not carry.
  if (
    op.revision === undefined &&
    op.at.type === "after" &&
    following?.type === "paragraph" &&
    location.paragraph.pPrMark?.survivor === "first"
  ) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "Insertion after a pending paragraph mark needs its former formatting snapshot and revision ids.",
    );
  }
  // Validate incoming identities and model records before creating any review records.
  const direct = applyOps(document, [replacement(canonicalBlocks)]);
  if (direct.isErr())
    return Result.err(
      new DocumentOpRefusal({
        reason: direct.error.reason,
        message: direct.error.message,
        opType: op.type,
      }),
    );
  const insertedIds = new Set(
    canonicalBlocks.map(({ paraId }) => idKey(paraId ?? "")),
  );
  const finalMarks = storyParagraphs(
    storyBody(direct.value.document, op.story),
  ).some(
    (at) =>
      insertedIds.has(idKey(at.paragraph.paraId ?? "")) &&
      hasContainerFinalParagraphMark(
        storyBody(direct.value.document, op.story),
        at,
      ),
  );
  if (finalMarks) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
      "Insertion would leave a tracked mark on a container-final paragraph.",
    );
  }
  const touched = {
    // The anchor's mark becomes the inserted one when it takes over.
    modified:
      takesOverMark && location.paragraph.paraId !== undefined
        ? [location.paragraph.paraId]
        : [],
    inserted: op.blocks.flatMap(({ paraId }) =>
      paraId === undefined ? [] : [paraId],
    ),
    removed: [],
  };
  if (op.revision === undefined) return Result.ok({ ...direct.value, touched });
  if (identityKeysIn(op.blocks).length > 0) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "Tracked insertion of identified review or content-control records is unsupported.",
    );
  }
  const revision = op.revision;
  // A move's destination marks its paragraphs moved to, not inserted.
  const markKind = op.move === true ? "moveTo" : "ins";
  const blocks: Paragraph[] = [];
  for (const [index, paragraph] of canonicalBlocks.entries()) {
    const wrapped = insertedContent({
      paragraph,
      revision: op.revision,
      move: op.move,
    });
    if (wrapped.isErr()) return wrapped;
    const content = wrapped.value;
    if (!takesOverMark || index < canonicalBlocks.length - 1) {
      blocks.push({
        ...paragraph,
        content,
        pPrMark: {
          kind: markKind,
          info: stampInfo(revision),
          // After a final paragraph each new break is rejected into the
          // paragraph before it, which ends up being the anchor, as the
          // anchor's own new break is.
          ...(takesOverMark ? { survivor: "first" as const } : {}),
        },
      });
      continue;
    }
    // The anchor's former mark: its properties before are the anchor's.
    const previous = location.paragraph.formatting;
    const current = finalizeParagraphFormatting(paragraph.formatting);
    blocks.push({
      ...paragraph,
      content,
      ...(sameFormatting(previous, current)
        ? {}
        : {
            propertyChanges: [
              {
                type: "paragraphPropertyChange",
                info: stampInfo(revision),
                ...(previous === undefined
                  ? {}
                  : { previousFormatting: previous }),
              },
            ],
          }),
    });
  }
  const anchorBlocks: Paragraph[] = takesOverMark
    ? [
        {
          ...location.paragraph,
          pPrMark: {
            kind: markKind,
            info: stampInfo(revision),
            survivor: "first",
          },
        },
      ]
    : [];
  // The anchor's own records are carried over and keep their ids; only the
  // records this insertion creates take new ones.
  const fresh = freshenIdentities({
    before: takesOverMark ? [location.paragraph] : [],
    after: [...anchorBlocks, ...blocks],
    newIds: op.newIds ?? {},
    usedElsewhere: () => new Set(packageIdentityKeys(document.package)),
  });
  switch (fresh.kind) {
    case "needsIds":
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        "Inserted paragraph review records need more revision ids.",
      );
    case "invalidId":
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        "An inserted review record has an invalid revision id.",
      );
    case "fresh": {
      const [replacedAnchor, ...inserted] = takesOverMark
        ? fresh.paragraphs
        : [anchor, ...fresh.paragraphs];
      const tracked = applyOps(document, [
        replacement(inserted, replacedAnchor),
      ]);
      return tracked.isErr()
        ? Result.err(
            new DocumentOpRefusal({
              reason: tracked.error.reason,
              message: tracked.error.message,
              opType: op.type,
            }),
          )
        : Result.ok({ ...tracked.value, touched });
    }
    default: {
      const unreachable: never = fresh;
      return unreachable;
    }
  }
};
