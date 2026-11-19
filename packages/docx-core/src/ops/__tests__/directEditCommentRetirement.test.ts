/**
 * A direct edit that removes every character a comment covers retires the
 * comment's thread, as a direct block deletion does: the save drops an
 * exhausted thread, so the model must not keep showing it.
 */
import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { allocateEditorIntentIds, compileEditorIntent } from "../editorIntent";
import { OP_STORIES } from "../types";

const PARA_ID = "12345678";
const DATE = "2026-01-01T00:00:00Z";

const run = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const commented = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: PARA_ID,
          content: [
            run("Keep "),
            { type: "commentRangeStart", id: 1 },
            run("gone"),
            { type: "commentRangeEnd", id: 1 },
            { type: "commentReference", id: 1 },
            run(" keep."),
          ],
        },
      ],
      comments: [
        {
          id: 1,
          author: "A",
          date: DATE,
          anchorKind: "range",
          content: [
            {
              type: "paragraph",
              paraId: "0A000001",
              content: [run("Note.")],
            },
          ],
        },
        {
          id: 2,
          parentId: 1,
          author: "B",
          date: DATE,
          content: [
            {
              type: "paragraph",
              paraId: "0A000002",
              content: [run("Reply.")],
            },
          ],
        },
      ],
    },
  },
});

const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: PARA_ID,
  offset,
});

const replace = (
  document: Document,
  from: number,
  to: number,
  text: string,
  tracked = false,
) => {
  const intent = {
    type: "replaceText" as const,
    from: at(from),
    to: at(to),
    text,
  };
  const ids = allocateEditorIntentIds(document, intent);
  const compiled = compileEditorIntent(document, {
    intent,
    mode: tracked
      ? {
          type: "suggesting",
          revision: { id: ids.revisionId, author: "A", date: DATE },
          newIds: ids.newIds,
        }
      : { type: "editing", newIds: ids.newIds },
  }).unwrap();
  return applyDocumentOps(document, compiled.ops).unwrap();
};

const markersOf = (document: Document): string[] => {
  const [block] = document.package.document.content;
  if (block?.type !== "paragraph") return [];
  return block.content.flatMap((node) =>
    node.type.startsWith("comment") ? [node.type] : [],
  );
};

test("deleting every commented character directly retires the thread", () => {
  const source = commented();
  const edited = replace(source, 5, 9, "");
  expect(edited.document.package.document.comments ?? []).toEqual([]);
  expect(markersOf(edited.document)).toEqual([]);
  // Undoing the edit brings the thread back as it was.
  const undone = applyDocumentOps(edited.document, edited.inverse).unwrap();
  expect(undone.document.package.document.comments).toEqual(
    source.package.document.comments,
  );
  expect(markersOf(undone.document)).toEqual(markersOf(source));
});

test("a direct edit that leaves commented characters keeps the thread", () => {
  const edited = replace(commented(), 5, 7, "");
  expect(
    (edited.document.package.document.comments ?? []).map(({ id }) => id),
  ).toEqual([1, 2]);
});

test("a tracked deletion of every commented character keeps the thread", () => {
  const edited = replace(commented(), 5, 9, "", true);
  expect(
    (edited.document.package.document.comments ?? []).map(({ id }) => id),
  ).toEqual([1, 2]);
});
