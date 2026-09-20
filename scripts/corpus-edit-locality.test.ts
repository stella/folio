import { FolioDocxReviewer } from "@stll/folio-core/ai-edits/headless";
import { buildParagraphsDocx } from "@stll/folio-core/ai-edits/__fixtures__/paragraphs";
import { buildBodySequenceDocx } from "@stll/folio-core/compare/__fixtures__/body-sequence";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { describe, expect, test } from "bun:test";

import type { CorpusInvariantInput } from "./lib/corpus-invariants/contract";
import {
  INSERTED_CHARACTER,
  firstEditableBlock,
  insertOneCharacterOperation,
  runEditLocalityInvariant,
} from "./lib/corpus-invariants/edit-locality";
import { generalizePartPath } from "./lib/corpus-invariants/save-idempotence";

const inputFor = async (buffer: ArrayBuffer): Promise<CorpusInvariantInput> => ({
  bytes: new Uint8Array(buffer),
  buffer,
  parsed: await parseDocx(buffer, { preloadFonts: false }),
  documentPart: "word/document.xml",
  budgetMs: 30_000,
});

const failuresFor = async (buffer: ArrayBuffer) =>
  (await runEditLocalityInvariant(await inputFor(buffer))).failures;

/** Run the invariant's own edit so a test can read where it landed. */
const editedBlockTexts = async (buffer: ArrayBuffer): Promise<string[]> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  const target = firstEditableBlock(reviewer.snapshot().blocks);
  if (target === undefined) {
    throw new Error("the fixture has no editable block");
  }
  reviewer.applyOperations([insertOneCharacterOperation(target)], { mode: "direct" });
  const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
  return reopened.snapshot().blocks.map((block) => block.text);
};

describe("edit locality", () => {
  test("a one-character edit reaches one paragraph and nothing else", async () => {
    const buffer = await buildBodySequenceDocx([
      { kind: "paragraph", text: "First paragraph." },
      { kind: "paragraph", text: "Second paragraph." },
    ]);

    expect(await editedBlockTexts(buffer)).toEqual([
      `First paragraph.${INSERTED_CHARACTER}`,
      "Second paragraph.",
    ]);
    expect(await failuresFor(buffer)).toEqual([]);
  });

  test("a package with a styles part and a header part keeps both", async () => {
    const buffer = await buildBodySequenceDocx(
      [
        { kind: "paragraph", text: "Body paragraph.", styleId: "Heading1" },
        { kind: "paragraph", text: "Another body paragraph." },
      ],
      { header: [{ kind: "paragraph", text: "Header paragraph." }] },
    );

    expect(await failuresFor(buffer)).toEqual([]);
  });

  test("a body with nothing to edit is not a defect", async () => {
    const buffer = await buildBodySequenceDocx([{ kind: "paragraph", text: "" }]);

    expect(await failuresFor(buffer)).toEqual([]);
  });

  /**
   * A package whose paragraphs carry no `w14:paraId` is addressed by aligned
   * ordinal rather than by id, so the edit is spliced and no untouched
   * paragraph is re-serialized. The invariant used to report the whole part
   * changing here, and the minted ids being written to disk with it.
   */
  test("an id-less package keeps the paragraphs the edit did not touch", async () => {
    const buffer = await buildParagraphsDocx(["First paragraph.", "Second paragraph."]);

    expect(await failuresFor(buffer)).toEqual([]);
  });

  test("a part path loses its per-package numbers", () => {
    expect(generalizePartPath("word/header12.xml")).toBe("word/headerN.xml");
  });
});
