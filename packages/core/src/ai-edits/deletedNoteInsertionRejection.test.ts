import { panic } from "better-result";
import { expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { type FolioDocumentStoryHandle, FolioDocxReviewer } from "./headless";

type NoteStory = Extract<
  FolioDocumentStoryHandle,
  { type: "footnote" | "endnote" }
>;
const INSERTION_FORMATS = {
  paragraph: {},
  heading: { styleId: "Heading2" },
  list: { numbering: { kind: "newList", format: "numbered" } },
} as const;

const shape = (reviewer: FolioDocxReviewer, story: NoteStory) => {
  const snapshot = reviewer.snapshotStory(story);
  if (!snapshot) panic("Generated note story disappeared", { story });
  return snapshot.blocks.map((block) => {
    if (block.kind === "diagnostic")
      panic("Unexpected diagnostic in generated note");
    const {
      kind,
      text,
      styleId,
      headingLevel,
      statedNumbering,
      displayLabel,
      directAlignment,
    } = block;
    return {
      kind,
      text,
      styleId,
      headingLevel,
      statedNumbering,
      displayLabel,
      directAlignment,
    };
  });
};

// The original generator edited live notes. Cross reference deletion with
// appended paragraphs, so occupied boundary carriers are exercised every run.
for (const noteType of ["footnote", "endnote"] as const) {
  for (const [insertionKind, formatting] of Object.entries(INSERTION_FORMATS)) {
    for (const paragraphs of [1, 2]) {
      for (const session of ["live", "reopened"] as const) {
        test(`${noteType}: reject ${paragraphs} appended ${insertionKind} paragraphs / ${session}`, async () => {
          const document = fromMarkdown("Body.");
          document.package.document.content.push(
            paragraph([
              run("Reference"),
              {
                type: "run",
                content: [
                  {
                    type:
                      noteType === "footnote" ? "footnoteRef" : "endnoteRef",
                    id: 37,
                  },
                ],
              },
            ]),
          );
          const content = [
            paragraph("Original note."),
            paragraph("Note closing."),
          ];
          if (noteType === "footnote")
            document.package.footnotes = [{ type: noteType, id: 37, content }];
          else
            document.package.endnotes = [{ type: noteType, id: 37, content }];
          let reviewer = await FolioDocxReviewer.fromBuffer(
            await createDocx(document),
          );
          const story = { type: noteType, noteId: 37 };
          const before = shape(reviewer, story);
          const reference = reviewer.getContent().at(-1);
          const anchor = reviewer.snapshotStory(story)?.blocks.at(-1);
          if (!reference || !anchor)
            panic("Generated note fixture lost its anchor");
          const deleted = reviewer.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode: "tracked-changes",
              operations: [
                { id: "delete", type: "deleteBlock", blockId: reference.id },
              ],
            },
            { undefinedReferences: "refuse" },
          );
          expect(deleted.applied).toHaveLength(1);
          const result = reviewer.applyDocumentOperationsToStory({
            story,
            batch: {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode: "tracked-changes",
              operations: [
                {
                  id: "insert",
                  type: "insertAfterBlock",
                  blockId: anchor.id,
                  text: Array.from(
                    { length: paragraphs },
                    (_, index) => `Inserted ${index}.`,
                  ).join("\n"),
                  formattingScope: "allParagraphs",
                  ...formatting,
                },
              ],
            },
            undefinedReferences: "refuse",
          });
          expect(result.applied).toHaveLength(1);
          if (session === "reopened")
            reviewer = await FolioDocxReviewer.fromBuffer(
              await reviewer.toBuffer(),
            );
          const originalView = reviewer.readReviewedStory({
            story,
            view: "original",
          });
          expect(
            originalView?.snapshot.blocks.map(({ kind, text }) => ({
              kind,
              text,
            })),
          ).toEqual(before.map(({ kind, text }) => ({ kind, text })));
          expect(reviewer.rejectAll()).toBeGreaterThan(0);
          expect(shape(reviewer, story)).toEqual(before);
          const reopened = await FolioDocxReviewer.fromBuffer(
            await reviewer.toBuffer(),
          );
          expect(shape(reopened, story)).toEqual(before);
        });
      }
    }
  }
}
