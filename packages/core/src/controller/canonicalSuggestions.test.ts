import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { getSuggestions } from "../prosemirror/commands/comments";
import { schema } from "../prosemirror/schema";
import type { Document, ParagraphContent } from "../types/document";

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

describe("resolving a pending suggestion in the canonical session", () => {
  const suggested = () =>
    createCanonicalSession(
      documentOf([
        run("Hello"),
        { type: "insertion", info: info(1, "s1"), content: [run(" there")] },
        { type: "insertion", info: info(3), content: [run("!")] },
      ]),
    ).unwrap();
  const infosOf = (session: ReturnType<typeof suggested>) => {
    const [paragraph] = session.document.package.document.content;
    return paragraph?.type === "paragraph"
      ? paragraph.content.flatMap((item) =>
          item.type === "insertion" ? [item.info] : [],
        )
      : [];
  };
  const resolve = (
    session: ReturnType<typeof suggested>,
    resolution: "accept" | "reject",
  ) => {
    const state = EditorState.create({ schema, doc: session.projection.doc });
    const commit = session
      .prepareSuggestion(state, {
        suggestionId: "s1",
        resolution,
        author: "Adopter",
        date: "2026-02-02T00:00:00Z",
      })
      .unwrap();
    expect(commit.publish().isOk()).toBe(true);
    return state.apply(commit.transaction);
  };

  test("accepting adopts it as an ordinary tracked change, and undoes", () => {
    const session = suggested();
    const before = session.document;
    const after = resolve(session, "accept");
    expect(infosOf(session)).toEqual([
      { id: 1, author: "Adopter", date: "2026-02-02T00:00:00Z" },
      info(3),
    ]);
    expect(
      getSuggestions(
        EditorState.create({ schema, doc: session.projection.doc }),
      ),
    ).toEqual([]);
    expect(session.prepareUndo(after).unwrap().publish().isOk()).toBe(true);
    expect(session.document).toEqual(before);
  });

  test("rejecting removes its revisions and keeps ordinary tracked changes", () => {
    const session = suggested();
    resolve(session, "reject");
    expect(infosOf(session)).toEqual([info(3)]);
  });
});
