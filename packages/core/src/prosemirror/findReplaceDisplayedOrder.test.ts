import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import type { Paragraph, Run } from "../types/content";
import type { Document } from "../types/document";
import { createDefaultFindOptions } from "../utils/findReplace";
import { toProseDoc } from "./conversion/toProseDoc";
import { findInDisplayedDocument } from "./findReplaceSelection";

const run = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });
const paragraph = (text: string): Paragraph => ({ type: "paragraph", content: [run(text)] });

const document: Document = {
  package: {
    document: {
      content: [
        paragraph("before hit"),
        {
          type: "paragraph",
          content: [
            run("anchor hit "),
            {
              type: "run",
              content: [
                {
                  type: "shape",
                  shape: {
                    type: "shape",
                    shapeType: "textBox",
                    size: { width: 914400, height: 914400 },
                    textBody: { content: [paragraph("box hit")] },
                  },
                },
              ],
            },
          ],
        },
        {
          type: "paragraph",
          content: [
            { type: "deletion", info: { id: 1, author: "Author" }, content: [run("deleted hit ")] },
            {
              type: "insertion",
              info: { id: 2, author: "Author" },
              content: [run("inserted hit ")],
            },
            run("after hit"),
          ],
        },
      ],
    },
  },
};

describe("search in a resolved review view", () => {
  test.each(["no-markup", "original"] as const)(
    "%s visits text-box matches after their anchor paragraph and before later body text",
    (displayMode) => {
      const state = EditorState.create({ doc: toProseDoc(document) });
      const matches = findInDisplayedDocument({
        document,
        state,
        searchText: "hit",
        options: { ...createDefaultFindOptions(), displayMode },
      });
      expect(
        matches?.map(({ paragraphIndex, textBoxIndex }) => ({ paragraphIndex, textBoxIndex })),
      ).toEqual([
        { paragraphIndex: 0, textBoxIndex: undefined },
        { paragraphIndex: 1, textBoxIndex: undefined },
        { paragraphIndex: 0, textBoxIndex: 0 },
        { paragraphIndex: 2, textBoxIndex: undefined },
        { paragraphIndex: 2, textBoxIndex: undefined },
      ]);
      expect(matches?.find((match) => match.textBoxIndex === 0)).toMatchObject({
        anchorParagraphIndex: 1,
      });
    },
  );
});
