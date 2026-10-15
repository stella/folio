import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";

import { createDefaultFindOptions } from "../utils/findReplace";
import { findInMarkupView } from "./findReplaceSelection";
import { schema } from "./schema";

const options = createDefaultFindOptions();
const revisionMark = (
  type: "insertion" | "deletion",
  moveKind: "moveFrom" | "moveTo" | null = null,
) =>
  (schema.marks[type] ?? panic(`Missing ${type} mark`)).create({
    revisionId: type === "insertion" ? 1 : 2,
    author: "Reviewer",
    date: "2026-01-01T00:00:00Z",
    moveKind,
  });

const doc = schema.node("doc", null, [
  schema.node("paragraph", null, [
    schema.text("old", [revisionMark("deletion")]),
    schema.text("new", [revisionMark("insertion")]),
    schema.text(" tail"),
  ]),
  schema.node("paragraph", null, [
    schema.text("moved", [revisionMark("deletion", "moveFrom")]),
    schema.text(" destination "),
    schema.text("moved", [revisionMark("insertion", "moveTo")]),
  ]),
]);
const state = EditorState.create({ doc });

describe("search in a markup view", () => {
  test("searches only displayed insertion, deletion and move text", () => {
    for (const [markupView, expectedOld, expectedNew, expectedMoves] of [
      ["all-markup", 1, 1, 2],
      ["no-markup", 0, 1, 1],
      ["simple-markup", 0, 1, 1],
      ["original", 1, 0, 1],
    ] as const) {
      const find = (searchText: string) =>
        findInMarkupView({ state, markupView, searchText, options });
      expect(find("old")).toHaveLength(expectedOld);
      expect(find("new")).toHaveLength(expectedNew);
      expect(find("moved")).toHaveLength(expectedMoves);
      for (const match of find("moved")) {
        expect(doc.textBetween(match.from, match.to)).toBe("moved");
      }
    }
  });

  test("maps a match at a hidden prefix boundary to the visible authored characters", () => {
    const finalMatch = findInMarkupView({
      state,
      markupView: "no-markup",
      searchText: "new tail",
      options,
    }).at(0);
    expect(finalMatch).toMatchObject({ from: 4, to: 12, text: "new tail" });
    const originalMatch = findInMarkupView({
      state,
      markupView: "original",
      searchText: "tail",
      options,
    }).at(0);
    expect(originalMatch).toMatchObject({ from: 8, to: 12, text: "tail" });
  });

  test("finds text joined by hiding a revision, preserving the authored spanning range", () => {
    const joined = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("left"),
        schema.text(" hidden ", [revisionMark("deletion")]),
        schema.text("right"),
      ]),
    ]);
    const joinedState = EditorState.create({ doc: joined });
    expect(
      findInMarkupView({
        state: joinedState,
        markupView: "no-markup",
        searchText: "leftright",
        options,
      }),
    ).toMatchObject([{ from: 1, to: 18, text: "leftright" }]);
    expect(
      findInMarkupView({
        state: joinedState,
        markupView: "all-markup",
        searchText: "leftright",
        options,
      }),
    ).toEqual([]);
    expect(joinedState.doc).toBe(joined);
  });
});
