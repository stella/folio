import { panic } from "better-result";
import { expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const open = (bytes: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes, {
    author: "Reviewer",
    session: "canonical",
  });

const rows = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ id, text }) => ({ id, text }));

const edit = (
  reviewer: FolioDocxReviewer,
  mode: FolioDocumentOperationMode,
  operation: FolioDocumentOperation,
) => {
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [operation],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.applied).toHaveLength(1);
};

test.each([
  "reject all",
  "reject insertion mark first",
  "reject split mark first",
  "delete direct",
  "delete tracked",
  "join direct",
] as const)(
  "a split of an inserted paragraph preserves the following paragraph ID after %s",
  async (resolution) => {
    const reviewer = await open(
      await createDocx(fromMarkdown("Before.\n\nAfter.")),
    );
    const original = rows(reviewer);
    const before =
      original.at(0) ?? panic("The fixture has a first paragraph.");
    edit(reviewer, "tracked-changes", {
      id: "insert",
      type: "insertAfterBlock",
      blockId: before.id,
      text: "Added clause.",
    });
    const inserted =
      rows(reviewer).at(1) ?? panic("The inserted paragraph is visible.");
    edit(reviewer, "tracked-changes", {
      id: "split",
      type: "splitBlock",
      blockId: inserted.id,
      offset: 6,
    });
    const second =
      rows(reviewer).at(2) ?? panic("The split's second half is visible.");
    const reopened = await open(await reviewer.toBuffer());
    if (
      resolution === "reject insertion mark first" ||
      resolution === "reject split mark first"
    ) {
      const blockId =
        resolution === "reject split mark first" ? inserted.id : second.id;
      const mark =
        reopened
          .getChanges()
          .find(
            (change) =>
              change.type === "paragraphMarkInserted" &&
              change.blockId === blockId,
          ) ?? panic("The selected paragraph has a pending mark.");
      expect(reopened.rejectChange(mark)).toBe(true);
    } else if (
      resolution === "delete direct" ||
      resolution === "delete tracked"
    ) {
      edit(
        reopened,
        resolution === "delete direct" ? "direct" : "tracked-changes",
        {
          id: "delete",
          type: "deleteBlock",
          blockId: second.id,
        },
      );
    } else if (resolution === "join direct") {
      edit(reopened, "direct", {
        id: "join",
        type: "mergeBlockWithNext",
        blockId: second.id,
      });
    }
    reopened.rejectAll();
    expect(rows(reopened)).toEqual(original);
  },
);
