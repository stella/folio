/** Run a synthetic edit through the public saved-document operation contract. */

import { FolioDocxReviewer } from "../../packages/core/src/ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../../packages/core/src/document-operations";
import type { EditOperationScript } from "./edit-operation-scripts";

export const runEditOperationScript = async (
  source: ArrayBuffer,
  script: EditOperationScript,
): Promise<ArrayBuffer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source, { author: "Synthetic Reviewer" });
  const anchor = reviewer.getContent().find(({ text }) => text === script.anchorText);
  if (!anchor) throw new Error(`${script.id}: missing anchor ${script.anchorText}`);

  if (script.action === "replyToComment") {
    const comment = reviewer
      .getComments()
      .find(({ anchoredText }) => anchoredText.includes(script.anchorText));
    if (!comment || !reviewer.replyTo(comment, { text: script.text })) {
      throw new Error(`${script.id}: comment range was unavailable`);
    }
    return reviewer.toBuffer();
  }

  let operation: FolioDocumentOperation;
  switch (script.action) {
    case "deleteBlock":
      operation = { id: script.id, type: "deleteBlock", blockId: anchor.id };
      break;
    case "mergeBlockWithNext":
      operation = { id: script.id, type: "mergeBlockWithNext", blockId: anchor.id };
      break;
    case "insertAfterBlock":
      operation = {
        id: script.id,
        type: "insertAfterBlock",
        blockId: anchor.id,
        text: script.text,
      };
      break;
    case "deleteTableRow":
      operation = { id: script.id, type: "deleteTableRow", blockId: anchor.id };
      break;
    case "replaceInBlock":
      operation = {
        id: script.id,
        type: "replaceInBlock",
        blockId: anchor.id,
        find: script.find,
        replace: script.replace,
      };
      break;
    default: {
      const unreachable: never = script;
      throw new Error(`Unsupported script: ${unreachable}`);
    }
  }

  const result = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: script.mode,
    operations: [operation],
  });
  if (result.status !== "committed" || result.issues.length > 0) {
    throw new Error(`${script.id}: ${JSON.stringify(result)}`);
  }
  if ("resolution" in script && script.resolution === "acceptAll") {
    if (reviewer.acceptAll() === 0) {
      throw new Error(`${script.id}: no change to accept`);
    }
  }
  return reviewer.toBuffer();
};
