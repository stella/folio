import assert from "node:assert/strict";

import type { FolioDocumentOperation } from "@stll/folio-core/server";

/** Plain-text effects over the pinned paragraph fixture; no editor operations are applied. */
export const modelPendingTextEffect = (
  operation: FolioDocumentOperation,
  textByBlock: Map<string, string>,
) => {
  const blockId =
    operation.type === "replaceRange" ||
    operation.type === "formatRange" ||
    operation.type === "commentOnRange"
      ? operation.range.blockId
      : operation.blockId;
  const text = textByBlock.get(blockId);
  assert.ok(text !== undefined, `Text model has no block ${blockId}`);
  switch (operation.type) {
    case "replaceRange": {
      const { startOffset, endOffset } = operation.range;
      assert.ok(startOffset >= 0 && endOffset >= startOffset && endOffset <= text.length);
      textByBlock.set(
        blockId,
        text.slice(0, startOffset) + operation.replace + text.slice(endOffset),
      );
      return;
    }
    case "replaceInBlock": {
      const start = text.indexOf(operation.find);
      assert.ok(start >= 0);
      assert.equal(start, text.lastIndexOf(operation.find), "Pinned replacement must be unique");
      textByBlock.set(
        blockId,
        text.slice(0, start) + operation.replace + text.slice(start + operation.find.length),
      );
      return;
    }
    case "replaceBlock":
      textByBlock.set(blockId, operation.text);
      return;
    case "deleteBlock":
      textByBlock.delete(blockId);
      return;
    case "splitBlock": {
      const { offset, separator = "" } = operation;
      assert.ok(offset > 0 && offset < text.length);
      assert.equal(text.slice(offset, offset + separator.length), separator);
      const entries = [...textByBlock];
      const index = entries.findIndex(([id]) => id === blockId);
      entries.splice(
        index,
        1,
        [blockId, text.slice(0, offset)],
        [`model:${operation.id}:split`, text.slice(offset + separator.length)],
      );
      textByBlock.clear();
      for (const [id, value] of entries) textByBlock.set(id, value);
      return;
    }
    case "mergeBlockWithNext": {
      const entries = [...textByBlock];
      const index = entries.findIndex(([id]) => id === blockId);
      const next = entries.at(index + 1);
      assert.ok(next, "Pinned merge must have a following paragraph");
      textByBlock.set(blockId, text + (operation.separator ?? "") + next[1]);
      textByBlock.delete(next[0]);
      return;
    }
    case "insertBeforeBlock":
    case "insertAfterBlock": {
      const paragraphs =
        operation.lineBreakMode === "inline"
          ? [operation.text]
          : operation.text.split(/\r\n|[\r\n]/u).filter((paragraph) => paragraph.length > 0);
      const entries = [...textByBlock];
      const index = entries.findIndex(([id]) => id === blockId);
      entries.splice(
        index + (operation.type === "insertAfterBlock" ? 1 : 0),
        0,
        ...paragraphs.map((paragraph, ordinal): [string, string] => [
          `model:${operation.id}:${ordinal}`,
          paragraph,
        ]),
      );
      textByBlock.clear();
      for (const [id, value] of entries) textByBlock.set(id, value);
      return;
    }
    case "formatRange":
    case "commentOnRange":
    case "commentOnBlock":
    case "setBlockParagraphProperties":
      return;
    case "insertTable":
    case "deleteTable":
    case "insertSignatureTable":
    case "insertTableRow":
    case "deleteTableRow":
    case "insertTableColumn":
    case "deleteTableColumn":
    case "mergeTableCells":
    case "splitTableCell":
      assert.fail(`Pinned paragraph fixture cannot model table operation ${operation.type}`);
    default: {
      const unhandled: never = operation;
      assert.fail(`Unhandled proposal text effect ${JSON.stringify(unhandled)}`);
    }
  }
};
