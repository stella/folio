/**
 * Deleting the only paragraph of a note on the canonical session: the
 * paragraph keeps its mark and loses its words, tracked as directly, so the
 * note keeps the paragraph it must hold.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";

const story = { type: "footnote", noteId: 1 } as const;

const open = async () => {
  const document = fromMarkdown("Opening words.");
  document.package.footnotes = [
    { type: "footnote", id: 1, content: [paragraph("As defined.")] },
  ];
  document.package.document.content.push(
    paragraph([
      run("Goods are defined."),
      { type: "run", content: [{ type: "footnoteRef", id: 1 }] },
    ]),
  );
  return FolioDocxReviewer.fromBuffer(await createDocx(document), {
    author: "Editor",
    session: "canonical",
  });
};

const noteTexts = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshotStory(story)?.blocks.map(({ text }) => text);

test.each(["direct", "tracked-changes"] as const)(
  "the only paragraph of a note loses its words and stays (%s)",
  async (mode) => {
    const reviewer = await open();
    const [block] = reviewer.snapshotStory(story)?.blocks ?? [];
    if (block === undefined) throw new Error("The note has a paragraph.");
    const result = reviewer.applyDocumentOperationsToStory({
      story,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [{ id: "delete", type: "deleteBlock", blockId: block.id }],
      },
      undefinedReferences: "refuse",
    });
    expect(result.skipped).toEqual([]);
    expect(result.applied.map(({ id }) => id)).toEqual(["delete"]);
    if (mode === "tracked-changes") {
      reviewer.rejectAll();
      expect(noteTexts(reviewer)).toEqual(["As defined."]);
      return;
    }
    expect(noteTexts(reviewer)).toEqual([""]);
  },
);

test("accepting the tracked deletion leaves the note's paragraph empty", async () => {
  const reviewer = await open();
  const [block] = reviewer.snapshotStory(story)?.blocks ?? [];
  if (block === undefined) throw new Error("The note has a paragraph.");
  reviewer.applyDocumentOperationsToStory({
    story,
    batch: {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [{ id: "delete", type: "deleteBlock", blockId: block.id }],
    },
    undefinedReferences: "refuse",
  });
  reviewer.acceptAll();
  expect(noteTexts(reviewer)).toEqual([""]);
});
