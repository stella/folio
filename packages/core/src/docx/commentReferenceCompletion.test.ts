/**
 * Where a missing `w:commentReference` is filled in, and where it is not.
 *
 * `completeCommentReferences` runs from the ProseMirror conversion, so these
 * pin the boundary that placement depends on: an edit reaches it through that
 * conversion whichever surface drove the edit, markdown has no comment syntax
 * to arrive with a half-written comment through, and the save path leaves the
 * model as authored.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { toMarkdown } from "../markdown";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Comment, Document, ParagraphContent } from "../types/document";
import { createDocx } from "./rezip";

const COMMENT_ID = 0;

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const comment: Comment = {
  id: COMMENT_ID,
  author: "Dana Lindqvist",
  initials: "DL",
  date: "2024-01-01T00:00:00Z",
  content: [{ type: "paragraph", content: [run("Check this wording.")] }],
};

/** A range whose comment has no `commentReference` anywhere in the story. */
const referencelessDocument = (): Document => ({
  package: {
    document: {
      comments: [comment],
      content: [
        {
          type: "paragraph",
          paraId: "61000000",
          content: [
            { type: "commentRangeStart", id: COMMENT_ID },
            run("Commented body text."),
            { type: "commentRangeEnd", id: COMMENT_ID },
          ],
        },
        { type: "paragraph", paraId: "61000001", content: [run("Second body paragraph.")] },
      ],
    },
  },
});

const documentXml = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(buffer)).file("word/document.xml")?.async("text")) ?? "";

const referenceCount = (xml: string): number => xml.split("<w:commentReference").length - 1;

describe("completing a missing comment reference", () => {
  test("a headless edit gives the comment its mark back", async () => {
    const source = await createDocx(referencelessDocument());
    expect(referenceCount(await documentXml(source))).toBe(0);

    const reviewer = await FolioDocxReviewer.fromBuffer(source, { author: "Editor" });
    const target = reviewer.snapshot().blocks[1];
    expect(target?.text).toBe("Second body paragraph.");
    reviewer.applyOperations(
      [
        {
          id: "edit",
          type: "replaceInBlock",
          blockId: target?.id ?? "",
          find: "Second body paragraph.",
          replace: "Superseded wording.",
        },
      ],
      { mode: "direct" },
    );

    const saved = await documentXml(await reviewer.toBuffer());
    expect(referenceCount(saved)).toBe(1);
    // Word paints the mark where the reference sits, so it must follow the end.
    expect(saved.indexOf("<w:commentReference")).toBeGreaterThan(
      saved.indexOf("<w:commentRangeEnd"),
    );
  });

  // The markdown reader has no comment production at all, so an import cannot
  // arrive with a range at all, let alone one missing its reference.
  test("a markdown import carries no comment marker to complete", () => {
    const exported = toMarkdown(
      {
        package: {
          document: {
            comments: [comment],
            content: [
              {
                type: "paragraph",
                content: [
                  { type: "commentRangeStart", id: COMMENT_ID },
                  run("Commented body text."),
                  { type: "commentRangeEnd", id: COMMENT_ID },
                  { type: "commentReference", id: COMMENT_ID },
                ],
              },
            ],
          },
        },
      },
      { annotations: "html" },
    );

    const imported = fromMarkdown(exported);
    expect(imported.package.document.comments).toBeUndefined();
    expect(JSON.stringify(imported.package)).not.toContain("commentRange");
  });

  // The selective save keeps every paragraph an edit did not touch
  // byte-for-byte, so a reference the save path minted would live in the model
  // and never reach the package. Completion belongs to the conversion that
  // knows where the reference goes, not to serialization.
  test("the save path leaves an authored model as it found it", async () => {
    const authored = referencelessDocument();
    const saved = await documentXml(await createDocx(authored));

    expect(referenceCount(saved)).toBe(0);
    expect(saved).toContain('<w:commentRangeStart w:id="0"/>');
    expect(saved).toContain('<w:commentRangeEnd w:id="0"/>');
  });
});
