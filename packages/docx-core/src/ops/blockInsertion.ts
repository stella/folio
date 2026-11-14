/** Paragraph insertion, isolated from the inline and table operation implementations. */
import { Result } from "better-result";

import {
  type Document,
  finalizeParagraphFormatting,
  type Paragraph,
} from "../model/document";
import {
  blockListAt,
  endsItsContainer,
  storyBody,
  storyParagraphs,
} from "./blocks";
import type { DocumentEdit } from "./edits";
import { structurallyEqual } from "./equality";
import { freshenIdentities } from "./identity";
import { identityKeysIn, idKey, packageIdentityKeys } from "./ids";
import { asMoveHalf } from "./moves";
import { paragraphLength } from "./offsets";
import { isBlockBoundary } from "./rangeBoundaries";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { ApplyOps } from "./resolve";
import { stampInfo, WRAP_KINDS, wrapTracked } from "./review";
import { DOCUMENT_OP_TYPES, type InsertBlocksOp } from "./types";

const sameFormatting = (
  left: Paragraph["formatting"],
  right: Paragraph["formatting"],
): boolean =>
  structurallyEqual(
    finalizeParagraphFormatting(left) ?? {},
    finalizeParagraphFormatting(right) ?? {},
  );

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
  if (op.blocks.length === 0) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
      "No paragraphs to insert.",
    );
  }
  if (op.at.type !== "before" && op.at.type !== "after") {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "Insert before or after a paragraph.",
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
  const following = blockListAt(body.content, location.list)[
    location.index + 1
  ];
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
  // Before a table the new paragraph's own mark can be the inserted one, as
  // the legacy editor records it, since resolving it runs into the table.
  // At the end of the list, before a range boundary between blocks that no
  // resolution may carry content across, or before any other block, such as
  // a content control, there is no mark after the new paragraph's for it to
  // resolve against, so the anchor's mark is the inserted one.
  const takesOverMark =
    op.revision !== undefined &&
    !before &&
    (following === undefined ||
      isBlockBoundary(following) ||
      following.type !== "table");
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
      at.paragraph.pPrMark !== undefined &&
      endsItsContainer(storyBody(direct.value.document, op.story), at),
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
    const wrapped = wrapTracked({
      items: paragraph.content,
      from: { offset: 0, zeroWidthBefore: 0 },
      to: {
        offset: paragraphLength(paragraph),
        zeroWidthBefore: Number.MAX_SAFE_INTEGER,
      },
      kind: WRAP_KINDS.INSERTION,
      stamp: op.revision,
    });
    if (wrapped.kind === "refused") {
      return refuse(wrapped.reason, "The inserted content cannot be tracked.");
    }
    const wrappedContent =
      wrapped.kind === "wrapped" ? wrapped.content : paragraph.content;
    const content =
      op.move === true
        ? asMoveHalf(wrappedContent, revision.id, "to")
        : wrappedContent;
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
