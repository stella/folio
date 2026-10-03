import { replaceStoryBody } from "./stories";
/** Whole-table edits and their exact structural inverse. */
import { Result, panic } from "better-result";

import type { BlockContent, Document, DocumentBody, Paragraph, Table } from "../model/document";
import { hasIllegalXmlCharacters } from "../serialize/xmlEscape";
import {
  blockListAt,
  endsItsContainer,
  sameBlockList,
  storyBody,
  storyParagraphs,
  updateBlockList,
  withBodyContent,
  type ParagraphLocation,
} from "./blocks";
import { validateOpsDocument } from "./contract";
import type { DocumentEdit } from "./edits";
import { equalForStaleness, structurallyEqual } from "./equality";
import {
  collides,
  countIds,
  countKeys,
  identityKeysIn,
  idKey,
  isParaId,
  packageIdentityKeys,
  packageParagraphIds,
  paragraphIdsIn,
} from "./ids";
import { textsIn } from "./leaves";
import { freshenIdentities } from "./identity";
import { stampInfo, WRAP_KINDS } from "./review";
import { permitsCellFinalMark, trackTableRows } from "./tableTracking";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { locateTableRow } from "./tableLocation";
import {
  DOCUMENT_OP_TYPES,
  type DeleteTableOp,
  type InsertTableOp,
  type SetContainerBlocksOp,
} from "./types";

type TableOp = InsertTableOp | DeleteTableOp | SetContainerBlocksOp;
type RefusalOptions = { op: TableOp; reason: DocumentOpRefusal["reason"]; message: string };
const refuse = ({ op, reason, message }: RefusalOptions) =>
  Result.err(new DocumentOpRefusal({ opType: op.type, reason, message }));

const paragraphsIn = (blocks: readonly BlockContent[]): Paragraph[] =>
  storyParagraphs({ content: [...blocks] }).map(({ paragraph }) => paragraph);

const endsInParagraph = (blocks: readonly BlockContent[]): boolean => {
  const last = blocks.at(-1);
  if (last?.type === "blockSdt" || last?.type === "blockCustomXml")
    return endsInParagraph(last.content);
  return last?.type === "paragraph";
};

/** Tables and cells must have usable content and end in a paragraph. */
const validStructure = (blocks: readonly BlockContent[]): boolean =>
  blocks.every((block) => {
    switch (block.type) {
      case "table":
        return (
          block.rows.length > 0 &&
          block.rows.every(
            (row) =>
              row.cells.length > 0 &&
              row.cells.every(
                (cell) => endsInParagraph(cell.content) && validStructure(cell.content),
              ),
          )
        );
      case "blockSdt":
      case "blockCustomXml":
        return validStructure(block.content);
      case "paragraph":
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        return true;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  });

/** Exact inverses may restore cell marks after independently resolved rows. */
type StructuralCellMarkOptions = { body: DocumentBody; location: ParagraphLocation; op: TableOp };
const permitsStructuralCellMark = ({ body, location, op }: StructuralCellMarkOptions): boolean => {
  if (permitsCellFinalMark(body, location)) return true;
  const mark = location.paragraph.pPrMark;
  return (
    op.type === DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS &&
    location.list.some((step) => step.kind === "tableCell") &&
    (mark?.kind === "ins" || mark?.kind === "del")
  );
};

type CommitBlocksOptions = {
  document: Document;
  op: TableOp;
  anchor: ParagraphLocation;
  blocks: readonly BlockContent[];
};
const commitBlocks = ({
  document,
  op,
  anchor,
  blocks,
}: CommitBlocksOptions): Result<DocumentEdit, DocumentOpRefusal> => {
  const body = storyBody(document, op.story);
  const before = blockListAt(body.content, anchor.list);
  if (structurallyEqual(before, blocks)) {
    return Result.ok({
      document,
      inverse: [],
      touched: { modified: [], inserted: [], removed: [] },
    });
  }
  if (
    !blocks.some(
      (block) =>
        block.type === "paragraph" &&
        idKey(block.paraId ?? "") === idKey(anchor.paragraph.paraId ?? ""),
    )
  ) {
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: "The structural inverse requires a surviving paragraph in this block list.",
    });
  }
  // A table at a container's end needs a terminal carrier, constructed by a
  // separate operation; this direct primitive never invents a paragraph id.
  if (
    blocks.length === 0 ||
    !validStructure(before) ||
    !validStructure(blocks) ||
    (endsItsContainer(body, { list: anchor.list, index: before.length - 1 }) &&
      (!endsInParagraph(before) || !endsInParagraph(blocks)))
  ) {
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      message: "A table has rows and cells; each cell and the story retain a final paragraph.",
    });
  }
  const fixedBlocks = (items: readonly BlockContent[]) =>
    items.filter((block) => block.type !== "table" && block.type !== "paragraph");
  if (!structurallyEqual(fixedBlocks(before), fixedBlocks(blocks))) {
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: "A table structural inverse preserves surrounding block wrappers and markers.",
    });
  }
  const beforeParagraphs = paragraphsIn(before);
  const afterParagraphs = paragraphsIn(blocks);
  const beforeById = new Map(
    beforeParagraphs.map((paragraph) => [idKey(paragraph.paraId ?? ""), paragraph]),
  );
  const afterById = new Map(
    afterParagraphs.map((paragraph) => [idKey(paragraph.paraId ?? ""), paragraph]),
  );
  if (
    [...beforeParagraphs, ...afterParagraphs].some((paragraph) => {
      const id = idKey(paragraph.paraId ?? "");
      return !structurallyEqual(
        paragraph.sectionProperties,
        (beforeById.has(id) ? afterById : beforeById).get(id)?.sectionProperties,
      );
    })
  ) {
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
      message: "Table edits preserve section boundaries.",
    });
  }
  if (
    [...beforeParagraphs, ...afterParagraphs].some(
      ({ paraId }) =>
        paraId === undefined ||
        ((!beforeById.has(idKey(paraId)) || !afterById.has(idKey(paraId))) && !isParaId(paraId)),
    )
  ) {
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      message: "An added or removed paragraph needs a usable id.",
    });
  }
  if (
    [...beforeParagraphs, ...afterParagraphs].some(({ content }) =>
      textsIn(content).some(hasIllegalXmlCharacters),
    )
  ) {
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT,
      message: "The table holds text that cannot be written to XML.",
    });
  }
  // Removal must be invertible under the same final-mark rule as insertion.
  // Resolve a cell's terminal mark before removing the table carrying it.
  for (const location of storyParagraphs(body)) {
    const paragraph = location.paragraph;
    if (
      paragraph.pPrMark !== undefined &&
      beforeById.has(idKey(paragraph.paraId ?? "")) &&
      endsItsContainer(body, location) &&
      !permitsStructuralCellMark({ body, location, op }) &&
      !structurallyEqual(paragraph.pPrMark, afterById.get(idKey(paragraph.paraId ?? ""))?.pPrMark)
    ) {
      return refuse({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
        message: "Removing this final paragraph mark would make the structural inverse invalid.",
      });
    }
  }
  const remaining = countIds(packageParagraphIds(document.package));
  for (const id of paragraphIdsIn(before))
    remaining.set(idKey(id), (remaining.get(idKey(id)) ?? 1) - 1);
  if (collides(remaining, paragraphIdsIn(blocks))) {
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      message: "A paragraph id is already used in the package.",
    });
  }
  const records = countKeys(packageIdentityKeys(document.package));
  for (const key of identityKeysIn(before)) records.set(key, (records.get(key) ?? 1) - 1);
  const incoming = identityKeysIn(blocks);
  if (
    incoming.some((key) => (records.get(key) ?? 0) > 0) ||
    new Set(incoming).size !== incoming.length
  ) {
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      message: "A revision or content-control id is already used in the package.",
    });
  }
  // Clone incoming values once; unchanged siblings and paragraphs keep their
  // original references, including when an inverse was read back from JSON.
  const beforeBlocksById = new Map(
    before.flatMap((block) => {
      const id = paragraphIdsIn(block).at(0);
      return id === undefined ? [] : [[idKey(id), block] as const];
    }),
  );
  const reuseParagraphs = (items: readonly BlockContent[]): BlockContent[] =>
    items.map((block): BlockContent => {
      switch (block.type) {
        case "paragraph": {
          const known = beforeById.get(idKey(block.paraId ?? ""));
          return known !== undefined && structurallyEqual(known, block) ? known : block;
        }
        case "table":
          return {
            ...block,
            rows: block.rows.map((row) => ({
              ...row,
              cells: row.cells.map((cell) => ({ ...cell, content: reuseParagraphs(cell.content) })),
            })),
          };
        case "blockSdt":
        case "blockCustomXml":
          return { ...block, content: reuseParagraphs(block.content) };
        case "preservedBlock":
        case "bookmarkStart":
        case "bookmarkEnd":
          return block;
        default: {
          const unreachable: never = block;
          return unreachable;
        }
      }
    });
  const replacement = blocks.map((block) => {
    const id = paragraphIdsIn(block).at(0);
    const existing =
      id === undefined
        ? before.find((candidate) => structurallyEqual(candidate, block))
        : beforeBlocksById.get(idKey(id));
    if (existing !== undefined && structurallyEqual(existing, block)) return existing;
    const cloned = reuseParagraphs([structuredClone(block)]).at(0);
    if (cloned === undefined) return panic("Cloning one block must preserve one block.");
    return cloned;
  });
  const content = updateBlockList(body.content, anchor.list, () => replacement);
  const next = replaceStoryBody({
    document,
    story: op.story,
    body: withBodyContent(body, content),
  });
  const nextBody = storyBody(next, op.story);
  const previousParagraphs = new Map(
    storyParagraphs(body).map(({ paragraph }) => [idKey(paragraph.paraId ?? ""), paragraph]),
  );
  for (const location of storyParagraphs(nextBody)) {
    const paragraph = location.paragraph;
    if (
      paragraph.pPrMark !== undefined &&
      endsItsContainer(nextBody, location) &&
      !permitsStructuralCellMark({ body: nextBody, location, op }) &&
      !structurallyEqual(
        paragraph.pPrMark,
        previousParagraphs.get(idKey(paragraph.paraId ?? ""))?.pPrMark,
      )
    ) {
      return refuse({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
        message: "A table edit cannot add a final paragraph mark.",
      });
    }
  }
  const valid = validateOpsDocument(next);
  if (valid.isErr())
    return refuse({ op, reason: valid.error.reason, message: valid.error.message });
  return Result.ok({
    document: next,
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
        story: op.story,
        blockId: anchor.paragraph.paraId ?? "",
        expected: replacement,
        blocks: before,
      },
    ],
    touched: {
      modified: afterParagraphs
        .filter(
          (paragraph) =>
            beforeById.has(idKey(paragraph.paraId ?? "")) &&
            !structurallyEqual(paragraph, beforeById.get(idKey(paragraph.paraId ?? ""))),
        )
        .flatMap(({ paraId }) => (paraId === undefined ? [] : [paraId])),
      inserted: afterParagraphs
        .filter(({ paraId }) => !beforeById.has(idKey(paraId ?? "")))
        .flatMap(({ paraId }) => (paraId === undefined ? [] : [paraId])),
      removed: beforeParagraphs
        .filter(({ paraId }) => !afterById.has(idKey(paraId ?? "")))
        .flatMap(({ paraId }) => (paraId === undefined ? [] : [paraId])),
    },
  });
};

export const applyTableOp = (
  document: Document,
  op: TableOp,
): Result<DocumentEdit, DocumentOpRefusal> => {
  const body = storyBody(document, op.story);
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_TABLE: {
      if (op.at.type !== "before" && op.at.type !== "after")
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          message: "Insert before or after a paragraph.",
        });
      const anchor = storyParagraphs(body).find(
        ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(op.at.blockId),
      );
      if (anchor === undefined)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
          message: "The table insertion anchor does not exist.",
        });
      const blocks = [...blockListAt(body.content, anchor.list)];
      const terminal = op.at.type === "after" && endsItsContainer(body, anchor);
      if (op.terminal !== undefined && !terminal) {
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          message: "A terminal insertion must follow its container's final paragraph.",
        });
      }
      if (terminal && op.terminal === undefined) {
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          message: "A terminal table insertion supplies the preceding paragraph id.",
        });
      }
      let preceding: Paragraph | undefined;
      if (op.terminal !== undefined) {
        if (op.revision !== undefined && anchor.paragraph.pPrMark !== undefined) {
          return refuse({
            op,
            reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
            message: "The final paragraph already carries a mark revision.",
          });
        }
        preceding = {
          type: "paragraph",
          paraId: op.terminal.beforeBlockId,
          ...(anchor.paragraph.formatting === undefined
            ? {}
            : { formatting: anchor.paragraph.formatting }),
          content: anchor.paragraph.content,
        };
        blocks.splice(anchor.index, 1, preceding, op.table, { ...anchor.paragraph, content: [] });
      } else {
        blocks.splice(anchor.index + (op.at.type === "after" ? 1 : 0), 0, op.table);
      }
      if (op.revision === undefined) return commitBlocks({ document, op, anchor, blocks });
      if (
        op.table.preserved !== undefined ||
        op.table.bookmarks !== undefined ||
        op.table.carrierStack !== undefined ||
        (op.table.propertyChanges?.length ?? 0) > 0
      ) {
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          message:
            "Tracked table construction does not support captured markup or table property revisions.",
        });
      }
      const tracked = trackTableRows({
        document,
        op,
        rows: op.table.rows,
        revision: op.revision,
        newIds: op.newIds,
        kind: WRAP_KINDS.INSERTION,
      });
      if (tracked.isErr()) return Result.err(tracked.error);
      const direct = commitBlocks({ document, op, anchor, blocks });
      if (direct.isErr()) return direct;
      const table: Table = { ...op.table, rows: tracked.value };
      const tableIndex = anchor.index + (op.at.type === "after" ? 1 : 0);
      blocks[tableIndex] = table;
      if (preceding !== undefined) {
        const outside = new Set(packageIdentityKeys(document.package));
        for (const key of identityKeysIn(anchor.paragraph)) outside.delete(key);
        for (const key of identityKeysIn(table)) outside.add(key);
        const fresh = freshenIdentities({
          before: [anchor.paragraph],
          after: [{ ...preceding, pPrMark: { kind: "ins", info: stampInfo(op.revision) } }],
          newIds: op.newIds ?? {},
          usedElsewhere: () => outside,
        });
        switch (fresh.kind) {
          case "needsIds":
            return refuse({
              op,
              reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
              message: "The preceding table break needs an additional revision id.",
            });
          case "invalidId":
            return refuse({
              op,
              reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
              message: "The preceding table break has an invalid revision id.",
            });
          case "fresh": {
            const paragraph = fresh.paragraphs.at(0);
            if (paragraph === undefined)
              return panic("A terminal table insertion must retain its preceding paragraph.");
            blocks[anchor.index] = paragraph;
            break;
          }
          default: {
            const unreachable: never = fresh;
            return unreachable;
          }
        }
      }
      return commitBlocks({ document, op, anchor, blocks });
    }
    case DOCUMENT_OP_TYPES.DELETE_TABLE: {
      const located = locateTableRow(document, op);
      if (located.isErr()) return Result.err(located.error);
      const location = located.value;
      if (op.expected !== undefined && !equalForStaleness(location.table, op.expected))
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.STALE,
          message: "The table to remove has changed.",
        });
      const anchor = storyParagraphs(body).find((paragraph) =>
        sameBlockList(paragraph.list, location.list),
      );
      if (anchor === undefined)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          message: "Removing a table requires a surviving paragraph in its block list.",
        });
      const blocks = [...blockListAt(body.content, location.list)];
      if (op.revision === undefined) {
        blocks.splice(location.index, 1);
      } else {
        if (
          location.table.preserved !== undefined ||
          location.table.bookmarks !== undefined ||
          location.table.carrierStack !== undefined ||
          (location.table.propertyChanges?.length ?? 0) > 0
        ) {
          return refuse({
            op,
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            message:
              "Tracked table deletion does not support captured markup or table property revisions.",
          });
        }
        const tracked = trackTableRows({
          document,
          op,
          rows: location.table.rows,
          revision: op.revision,
          newIds: op.newIds,
          kind: WRAP_KINDS.DELETION,
        });
        if (tracked.isErr()) return Result.err(tracked.error);
        blocks[location.index] = { ...location.table, rows: tracked.value };
      }
      return commitBlocks({ document, op, anchor, blocks });
    }
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS: {
      const anchor = storyParagraphs(body).find(
        ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(op.blockId),
      );
      if (anchor === undefined)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
          message: "The structural inverse anchor does not exist.",
        });
      const before = blockListAt(body.content, anchor.list);
      if (!equalForStaleness(before, op.expected))
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.STALE,
          message: "The container blocks have changed.",
        });
      return commitBlocks({ document, op, anchor, blocks: op.blocks });
    }
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};
