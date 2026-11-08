import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { getSuggestions } from "../prosemirror/commands/comments";
import { schema } from "../prosemirror/schema";
import type { Document, ParagraphContent } from "../types/document";
import { createCanonicalSession } from "./canonicalSession";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const info = (id: number, suggestionId?: string) => ({
  id,
  author: "Reviewer",
  date: "2026-01-01T00:00:00Z",
  ...(suggestionId === undefined ? {} : { suggestionId }),
});
const documentOf = (content: ParagraphContent[]): Document => ({
  package: {
    document: {
      content: [{ type: "paragraph", paraId: "10000000", content }],
    },
  },
});

describe("pending suggestions in the canonical projection", () => {
  test("a suggestion's revisions show as that suggestion; a tracked change does not", () => {
    const session = createCanonicalSession(
      documentOf([
        run("Hello"),
        { type: "insertion", info: info(1, "s1"), content: [run(" there")] },
        { type: "deletion", info: info(2, "s1"), content: [run(" world")] },
        { type: "insertion", info: info(3), content: [run("!")] },
      ]),
    ).unwrap();
    const state = EditorState.create({ schema, doc: session.projection.doc });
    expect(
      getSuggestions(state).map(({ suggestionId }) => suggestionId),
    ).toEqual(["s1"]);
    const marks: { revisionId: unknown; provenance: unknown }[] = [];
    state.doc.descendants((node) => {
      for (const mark of node.marks)
        if (mark.type.name === "insertion" || mark.type.name === "deletion")
          marks.push({
            revisionId: mark.attrs["revisionId"],
            provenance: mark.attrs["provenance"],
          });
    });
    expect(marks).toEqual([
      { revisionId: 1, provenance: "suggested" },
      { revisionId: 2, provenance: "suggested" },
      { revisionId: 3, provenance: "user" },
    ]);
  });

  test("a suggested run or paragraph formatting change shows as its suggestion", () => {
    const session = createCanonicalSession({
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "10000000",
              formatting: { alignment: "center" },
              propertyChanges: [
                {
                  type: "paragraphPropertyChange",
                  info: info(5, "s2"),
                  previousFormatting: {},
                },
              ],
              content: [
                {
                  type: "run",
                  formatting: { bold: true },
                  propertyChanges: [
                    {
                      type: "runPropertyChange",
                      info: info(4, "s1"),
                      previousFormatting: {},
                    },
                  ],
                  content: [{ type: "text", text: "Hello" }],
                },
                {
                  type: "run",
                  formatting: { italic: true },
                  propertyChanges: [
                    {
                      type: "runPropertyChange",
                      info: info(6),
                      previousFormatting: {},
                    },
                  ],
                  content: [{ type: "text", text: " world" }],
                },
              ],
            },
          ],
        },
      },
    }).unwrap();
    const state = EditorState.create({ schema, doc: session.projection.doc });
    expect(
      getSuggestions(state)
        .map(({ suggestionId }) => suggestionId)
        .sort(),
    ).toEqual(["s1", "s2"]);
  });
});
