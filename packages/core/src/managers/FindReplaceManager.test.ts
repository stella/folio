import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { findInProseMirrorDocument } from "../prosemirror/findReplaceSelection";
import type { ComplexField, Run, SimpleField } from "../types/content";
import type { Document } from "../types/document";
import { resolveDocumentFindRanges } from "../utils/findMatchView";
import { createDefaultFindOptions, findInDocument } from "../utils/findReplace";
import type { ReplaceEligibility } from "../utils/findReplace";
import { FindReplaceManager, getAdjacentFindIndex } from "./FindReplaceManager";

type TestMatch = {
  paragraphIndex: number;
  startOffset: number;
  endOffset: number;
  id: string;
  replacement: ReplaceEligibility;
};

const match = (
  id: string,
  paragraphIndex: number,
  startOffset: number,
): TestMatch => ({
  id,
  paragraphIndex,
  startOffset,
  endOffset: startOffset + 1,
  replacement: { status: "allowed" },
});

const textRun = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});

const makeField = (kind: "simple" | "complex"): SimpleField | ComplexField => {
  if (kind === "simple")
    return {
      type: "simpleField",
      fieldType: "DOCVARIABLE",
      instruction: "DOCVARIABLE value",
      content: [textRun("hit hit")],
      fldLock: true,
    };
  return {
    type: "complexField",
    fieldType: "DOCVARIABLE",
    instruction: "DOCVARIABLE value",
    fieldCode: [
      {
        type: "run",
        content: [{ type: "instrText", text: "DOCVARIABLE value" }],
      },
    ],
    fieldResult: [textRun("hit hit")],
    fldLock: true,
  };
};

describe("find replacement eligibility", () => {
  const noteCases = (["footnoteRef", "endnoteRef"] as const).flatMap((type) =>
    [1, 12, 123].map((id) => ({ type, id })),
  );

  test.each(noteCases)(
    "$type id $id blocks replacement across its anchor",
    ({ type, id }) => {
      const source: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [
                  {
                    type: "run",
                    content: [
                      { type: "text", text: "A" },
                      { type, id },
                      { type: "text", text: "B" },
                    ],
                  },
                ],
              },
            ],
          },
        },
      };
      const matches = findInDocument(source, "AB", createDefaultFindOptions());
      expect(matches).toHaveLength(1);
      expect(matches.at(0)?.replacement).toEqual({
        status: "blocked",
        reason: "atomic-content",
      });
    },
  );

  test.each(noteCases)(
    "ordinary text after $type id $id remains replaceable in PM",
    ({ type, id }) => {
      const source: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [
                  {
                    type: "run",
                    content: [
                      { type, id },
                      { type: "text", text: " hit" },
                    ],
                  },
                ],
              },
            ],
          },
        },
      };
      const options = createDefaultFindOptions();
      const matches = findInDocument(source, "hit", options);
      expect(matches).toHaveLength(1);
      expect(matches.at(0)?.replacement).toEqual({ status: "allowed" });
      const pmDoc = toProseDoc(source);
      const pmMatches = findInProseMirrorDocument(pmDoc, "hit", options);
      expect(pmMatches).toHaveLength(1);
      const range = resolveDocumentFindRanges({
        document: source,
        bodyDoc: pmDoc,
        matches,
      }).at(0);
      if (!range) panic("Ordinary PM range after note anchor missing");
      expect(pmDoc.textBetween(range.from, range.to)).toBe("hit");
      expect(pmMatches.at(0)?.from).toBe(range.from);
      expect(pmMatches.at(0)?.to).toBe(range.to);
      const updatedPm = EditorState.create({ doc: pmDoc }).tr.insertText(
        "value",
        range.from,
        range.to,
      ).doc;
      expect(updatedPm.child(0).child(0).toJSON()).toEqual(
        pmDoc.child(0).child(0).toJSON(),
      );
      expect(findInProseMirrorDocument(updatedPm, "hit", options)).toHaveLength(
        0,
      );
      const updatedMatch = findInProseMirrorDocument(
        updatedPm,
        "value",
        options,
      ).at(0);
      if (!updatedMatch) panic("PM replacement after note anchor missing");
      expect(updatedPm.textBetween(updatedMatch.from, updatedMatch.to)).toBe(
        "value",
      );
    },
  );

  test.each(["simple", "complex"] as const)(
    "%s field substring hits are blocked from replacement",
    (kind) => {
      const source: Document = {
        package: {
          document: {
            content: [{ type: "paragraph", content: [makeField(kind)] }],
          },
        },
      };
      const matches = findInDocument(source, "it", createDefaultFindOptions());
      expect(matches.map(({ replacement }) => replacement)).toEqual([
        { status: "blocked", reason: "field-result" },
        { status: "blocked", reason: "field-result" },
      ]);
    },
  );

  test("a final-view match across hidden deletion text is blocked from replacement", () => {
    const source: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                textRun("A"),
                {
                  type: "deletion",
                  info: {
                    id: 17,
                    author: "Author",
                    date: "2026-01-01T00:00:00Z",
                  },
                  content: [textRun("X")],
                },
                textRun("B"),
              ],
            },
          ],
        },
      },
    };
    const matches = findInDocument(source, "AB", {
      ...createDefaultFindOptions(),
      displayMode: "no-markup",
    });
    expect(matches).toHaveLength(1);
    expect(matches.at(0)).toMatchObject({
      startOffset: 0,
      endOffset: 3,
      text: "AB",
    });
    expect(matches.at(0)?.replacement).toEqual({
      status: "blocked",
      reason: "hidden-content",
    });
  });
});

describe("getAdjacentFindIndex", () => {
  test("wraps forward past the last match to the first", () => {
    expect(getAdjacentFindIndex(2, 3, "next")).toBe(0);
  });

  test("wraps backward past the first match to the last", () => {
    expect(getAdjacentFindIndex(0, 3, "previous")).toBe(2);
  });

  test("steps within range", () => {
    expect(getAdjacentFindIndex(0, 3, "next")).toBe(1);
    expect(getAdjacentFindIndex(2, 3, "previous")).toBe(1);
  });

  test("returns 0 when there are no matches", () => {
    expect(getAdjacentFindIndex(0, 0, "next")).toBe(0);
    expect(getAdjacentFindIndex(0, 0, "previous")).toBe(0);
  });
});

describe("FindReplaceManager cursor", () => {
  test("setMatches seeds the cursor on the first match and keeps an empty result", () => {
    const manager = new FindReplaceManager<TestMatch>();
    const matches = [match("a", 0, 0), match("b", 0, 4)];

    expect(manager.setMatches(matches)).toEqual({
      matches,
      totalCount: 2,
      currentIndex: 0,
    });

    expect(manager.setMatches([])).toEqual({
      matches: [],
      totalCount: 0,
      currentIndex: 0,
    });
  });

  test("navigate wraps the cursor and reports the now-current match", () => {
    const manager = new FindReplaceManager<TestMatch>();
    const matches = [match("a", 0, 0), match("b", 1, 0)];
    manager.setMatches(matches);

    expect(manager.navigate("next")).toEqual({ match: matches[1]!, index: 1 });
    expect(manager.navigate("next")).toEqual({ match: matches[0]!, index: 0 });
    expect(manager.navigate("previous")).toEqual({
      match: matches[1]!,
      index: 1,
    });
    expect(manager.getResult()?.currentIndex).toBe(1);
  });

  test("navigate returns null with no matches and clear discards the result", () => {
    const manager = new FindReplaceManager<TestMatch>();
    expect(manager.navigate("next")).toBeNull();

    manager.setMatches([match("a", 0, 0)]);
    manager.clear();
    expect(manager.getResult()).toBeNull();
    expect(manager.navigate("next")).toBeNull();
  });

  test("keeps explicit jumps synchronized with subsequent navigation", () => {
    const manager = new FindReplaceManager<TestMatch>();
    const matches = [match("a", 0, 0), match("b", 0, 4), match("c", 0, 8)];
    manager.setMatches(matches);

    expect(manager.goTo(1)).toEqual({ match: matches[1]!, index: 1 });
    expect(manager.navigate("next")).toEqual({ match: matches[2]!, index: 2 });
    expect(manager.goTo(9)).toBeNull();
    expect(manager.getResult()?.currentIndex).toBe(2);
  });
});

// Refreshes keep the current cursor bounded as the result set shrinks.
test("result refresh preserves or clamps the cursor", () => {
  const manager = new FindReplaceManager<TestMatch>();
  const matches = [match("a", 0, 0), match("b", 1, 0), match("c", 2, 0)];
  for (const index of [-1, 0, 1, 2, 3, 20]) {
    const result = manager.setMatches(matches, index);
    expect(result.currentIndex).toBe(
      Math.max(0, Math.min(index, matches.length - 1)),
    );
    expect(manager.navigate("next")?.index).toBe(
      (result.currentIndex + 1) % matches.length,
    );
  }
  expect(manager.setMatches([], 4).currentIndex).toBe(0);
  expect(manager.navigate("next")).toBeNull();
});
