/**
 * Deleting a container's final paragraph retracts a pending inserted break
 * before it the same way before and after a save: the paragraph holding that
 * break reads its pending property change back from the package without the
 * values no package states, which must not read as a change of its own.
 */

import { describe, expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const open = async (bytes: ArrayBuffer | Uint8Array) =>
  FolioDocxReviewer.fromBuffer(bytes, { author: "Test" });

const apply = (
  reviewer: FolioDocxReviewer,
  operations: Parameters<
    FolioDocxReviewer["applyDocumentOperations"]
  >[0]["operations"],
) => {
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations,
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
};

/** A merge with merged properties, then a split of the paragraph it joins. */
const pending = async () => {
  const reviewer = await open(
    (
      await ensureParaIds(
        await createDocx(
          fromMarkdown("Opening.\n\nLaunch.\n\nOpens the party."),
        ),
      )
    ).docx,
  );
  const [, launch, closing] = reviewer.snapshot().blocks;
  if (!launch || !closing) throw new Error("Missing fixture");
  apply(reviewer, [
    {
      id: "merge",
      type: "mergeBlockWithNext",
      blockId: launch.id,
      separator: " ",
      mergedParagraphProperties: { alignment: "center" },
    },
    { id: "split", type: "splitBlock", blockId: closing.id, offset: 10 },
  ]);
  return reviewer;
};

const deleteFinalAndAccept = async (reviewer: FolioDocxReviewer) => {
  const final = reviewer.snapshot().blocks.at(-1);
  if (!final) throw new Error("No final paragraph");
  apply(reviewer, [{ id: "delete", type: "deleteBlock", blockId: final.id }]);
  reviewer.acceptAll();
  const reopened = await open(await reviewer.toBuffer());
  return reopened
    .toDocument()
    .package.document.content.map((block) =>
      block.type === "paragraph"
        ? block.content
            .flatMap((child) =>
              child.type === "run"
                ? child.content.flatMap((item) =>
                    item.type === "text" ? [item.text] : [],
                  )
                : [],
            )
            .join("")
        : block.type,
    );
};

describe("final paragraph deletion after a save", () => {
  test("retracts the inserted break before it as it does live", async () => {
    const live = await pending();
    const saved = await open(await (await pending()).toBuffer());
    const expected = ["Opening.", "Launch. Opens the "];
    expect(await deleteFinalAndAccept(live)).toEqual(expected);
    expect(await deleteFinalAndAccept(saved)).toEqual(expected);
  });
});
