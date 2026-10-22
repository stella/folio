import { panic } from "better-result";
import { expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph, run } from "../docx/server/build";
import { toMarkdown } from "../markdown";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { type FolioDocumentStoryHandle, FolioDocxReviewer } from "./headless";

const STORIES = {
  main: { type: "main" },
  header: { type: "header", relationshipId: "rIdHeader" },
  footer: { type: "footer", relationshipId: "rIdFooter" },
  footnote: { type: "footnote", noteId: 37 },
  endnote: { type: "endnote", noteId: 37 },
} as const satisfies Record<
  FolioDocumentStoryHandle["type"],
  FolioDocumentStoryHandle
>;

const AUTHORED_RUNS = {
  inherited: {},
  bold: { bold: true },
  italic: { italic: true },
} as const;

const markdownStory = (
  reviewer: FolioDocxReviewer,
  story: FolioDocumentStoryHandle,
) => {
  const document = reviewer.toDocument();
  const content = (() => {
    switch (story.type) {
      case "main":
        return document.package.document.content;
      case "header":
        return document.package.headers?.get(story.relationshipId)?.content;
      case "footer":
        return document.package.footers?.get(story.relationshipId)?.content;
      case "footnote":
        return document.package.footnotes?.find(({ id }) => id === story.noteId)
          ?.content;
      case "endnote":
        return document.package.endnotes?.find(({ id }) => id === story.noteId)
          ?.content;
      default: {
        const unreachable: never = story;
        return unreachable;
      }
    }
  })();
  if (!content) panic("Generated projection story disappeared", { story });
  document.package.document.content = content;
  return toMarkdown(document);
};

// Cross paragraph-style suggestions with authored run formatting in every
// story, so an inherited mark cannot escape through a secondary converter.
for (const story of Object.values(STORIES)) {
  for (const [runKind, formatting] of Object.entries(AUTHORED_RUNS)) {
    for (const direction of ["add", "remove"] as const) {
      test(`${story.type}: ${direction} suggested heading / ${runKind} runs share the saved projection`, async () => {
        const content = () => [
          paragraph(
            [run("Original clause.", formatting)],
            direction === "remove" ? { styleId: "Heading2" } : {},
          ),
        ];
        const document = fromMarkdown("Body.");
        document.package.document.content = content();
        document.package.document.content.push(
          paragraph([
            run("References"),
            {
              type: "run",
              content: [
                { type: "footnoteRef", id: 37 },
                { type: "endnoteRef", id: 37 },
              ],
            },
          ]),
        );
        document.package.headers = new Map([
          [
            STORIES.header.relationshipId,
            { type: "header", hdrFtrType: "default", content: content() },
          ],
        ]);
        document.package.footers = new Map([
          [
            STORIES.footer.relationshipId,
            { type: "footer", hdrFtrType: "default", content: content() },
          ],
        ]);
        document.package.footnotes = [
          { type: "footnote", id: 37, content: content() },
        ];
        document.package.endnotes = [
          { type: "endnote", id: 37, content: content() },
        ];
        document.package.document.finalSectionProperties = {
          ...document.package.document.finalSectionProperties,
          headerReferences: [
            { type: "default", rId: STORIES.header.relationshipId },
          ],
          footerReferences: [
            { type: "default", rId: STORIES.footer.relationshipId },
          ],
        };
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await createDocx(document),
        );
        const expected = markdownStory(reviewer, story);
        const fullMarkdown = toMarkdown(reviewer.toDocument());
        const notes = reviewer.getNotesAsText();
        const anchor = reviewer.snapshotStory(story)?.blocks.at(0);
        if (!anchor)
          panic("Generated projection story lost its anchor", { story });
        const result = reviewer.applyDocumentOperationsToStory({
          story,
          batch: {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "suggested",
            operations: [
              {
                id: "style",
                type: "setBlockParagraphProperties",
                blockId: anchor.id,
                properties: {
                  styleId: direction === "add" ? "Heading2" : null,
                },
              },
            ],
          },
          undefinedReferences: "refuse",
        });
        expect(result.applied).toHaveLength(1);
        const pending = reviewer.exportPendingSuggestions();
        expect(pending).toHaveLength(1);
        expect(markdownStory(reviewer, story)).toBe(expected);
        expect(toMarkdown(reviewer.toDocument())).toBe(fullMarkdown);
        expect(reviewer.getNotesAsText()).toBe(notes);
        const reopened = await FolioDocxReviewer.fromBuffer(
          await reviewer.toBuffer(),
        );
        expect(markdownStory(reopened, story)).toBe(expected);
        expect(toMarkdown(reopened.toDocument())).toBe(fullMarkdown);
        expect(reviewer.exportPendingSuggestions()).toEqual(pending);
      });
    }
  }
}
