import { describe, expect, test } from "bun:test";

import { FindReplaceManager, getAdjacentFindIndex } from "./FindReplaceManager";
import type { OpStory } from "@stll/docx-core/ops";
import type { Document } from "../types/document";
import type { ComplexField, Paragraph, Run, SimpleField } from "../types/content";
import { panic } from "better-result";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { findInProseMirrorDocument } from "../prosemirror/findReplaceSelection";
import { resolveDocumentFindRanges } from "../utils/findMatchView";
import { EditorState } from "prosemirror-state";
import { createDefaultFindOptions, findInDocument } from "../utils/findReplace";
import type { ReplaceEligibility } from "../utils/findReplace";

type TestMatch = {
  paragraphIndex: number;
  startOffset: number;
  endOffset: number;
  id: string;
  replacement: ReplaceEligibility;
};

const match = (id: string, paragraphIndex: number, startOffset: number): TestMatch => ({
  id,
  paragraphIndex,
  startOffset,
  endOffset: startOffset + 1,
  replacement: { status: "allowed" },
});

type StoryTestMatch = TestMatch & {
  story: OpStory;
  textBoxIndex?: number;
  commentId?: number;
};

const unsupportedMatches = [
  { ...match("header", 0, 0), story: { kind: "header", rId: "rId1" } },
  { ...match("footer", 0, 0), story: { kind: "footer", rId: "rId2" } },
  { ...match("footnote", 0, 0), story: { kind: "footnote", id: 1 } },
  { ...match("endnote", 0, 0), story: { kind: "endnote", id: 2 } },
  { ...match("shape", 0, 0), story: "main", textBoxIndex: 0 },
  { ...match("comment", 0, 0), story: "main", commentId: 7 },
] as const satisfies readonly StoryTestMatch[];

describe("FindReplaceManager replacement story boundaries", () => {
  test.each(unsupportedMatches)(
    "keeps $id navigable while refusing body replacement",
    (secondary) => {
      const document: Document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [{ type: "run", content: [{ type: "text", text: "body" }] }],
              },
            ],
          },
        },
      };
      const before = structuredClone(document);
      const body = { ...match("body", 0, 0), story: "main" } as const satisfies StoryTestMatch;
      const manager = new FindReplaceManager<StoryTestMatch>();
      manager.setMatches([body, secondary]);
      expect(manager.goTo(1)?.match).toEqual(secondary);
      expect(manager.replaceCurrent(document, "changed")).toBeNull();
      expect(manager.navigate("previous")?.match).toEqual(body);
      expect(manager.replaceCurrent(document, "B")).not.toBeNull();
      const replaced = manager.replaceAll(document, [body, secondary], "changed");
      expect(replaced).toMatchObject({ replacedCount: 1, skippedCount: 1 });
      expect(document).toEqual(before);
    },
  );
});

const textRun = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });

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
    fieldCode: [{ type: "run", content: [{ type: "instrText", text: "DOCVARIABLE value" }] }],
    fieldResult: [textRun("hit hit")],
    fldLock: true,
  };
};

const documentParagraph = (document: Document): Paragraph => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Fixture paragraph missing");
  return paragraph;
};

describe("FindReplaceManager real search replacement eligibility", () => {
  const noteCases = (["footnoteRef", "endnoteRef"] as const).flatMap((type) =>
    [1, 12, 123].map((id) => ({ type, id })),
  );

  test.each(noteCases)(
    "$type id $id blocks replacement across its anchor without changing source or PM",
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
      const before = structuredClone(source);
      const beforePm = toProseDoc(source).toJSON();
      const matches = findInDocument(source, "AB", createDefaultFindOptions());
      expect(matches).toHaveLength(1);
      expect(matches.at(0)?.replacement).toEqual({ status: "blocked", reason: "atomic-content" });
      const manager = new FindReplaceManager<(typeof matches)[number]>();
      manager.setMatches(matches);
      expect(manager.replaceCurrent(source, "value")).toBeNull();
      expect(manager.replaceAll(source, matches, "value")).toEqual({
        document: source,
        replacedCount: 0,
        skippedCount: 1,
      });
      expect(source).toEqual(before);
      expect(toProseDoc(source).toJSON()).toEqual(beforePm);
    },
  );

  test.each(noteCases)(
    "ordinary text after $type id $id remains replaceable in model and PM",
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
      const manager = new FindReplaceManager<(typeof matches)[number]>();
      manager.setMatches(matches);
      const replaced = manager.replaceCurrent(source, "value");
      if (!replaced) panic("Ordinary text replacement after note anchor failed");
      expect(findInDocument(replaced, "value", options)).toHaveLength(1);
      expect(documentParagraph(replaced).content.at(0)).toMatchObject({
        content: [
          { type, id },
          { type: "text", text: " " },
          { type: "text", text: "value" },
        ],
      });
      const pmDoc = toProseDoc(source);
      const pmMatches = findInProseMirrorDocument(pmDoc, "hit", options);
      expect(pmMatches).toHaveLength(1);
      const range = resolveDocumentFindRanges({ document: source, bodyDoc: pmDoc, matches }).at(0);
      if (!range) panic("Ordinary PM range after note anchor missing");
      expect(pmDoc.textBetween(range.from, range.to)).toBe("hit");
      expect(pmMatches.at(0)?.from).toBe(range.from);
      expect(pmMatches.at(0)?.to).toBe(range.to);
      const updatedPm = EditorState.create({ doc: pmDoc }).tr.insertText(
        "value",
        range.from,
        range.to,
      ).doc;
      expect(updatedPm.child(0).child(0).toJSON()).toEqual(pmDoc.child(0).child(0).toJSON());
      expect(findInProseMirrorDocument(updatedPm, "hit", options)).toHaveLength(0);
      const updatedMatch = findInProseMirrorDocument(updatedPm, "value", options).at(0);
      if (!updatedMatch) panic("PM replacement after note anchor missing");
      expect(updatedPm.textBetween(updatedMatch.from, updatedMatch.to)).toBe("value");
    },
  );

  test.each(["simple", "complex"] as const)(
    "%s field substring hits are skipped and preserve the field and PM document",
    (kind) => {
      const field = makeField(kind);
      const source: Document = {
        package: { document: { content: [{ type: "paragraph", content: [field] }] } },
      };
      const before = structuredClone(source);
      const beforePm = toProseDoc(source).toJSON();
      const matches = findInDocument(source, "it", createDefaultFindOptions());
      expect(matches).toHaveLength(2);
      expect(matches.map(({ replacement }) => replacement)).toEqual([
        { status: "blocked", reason: "field-result" },
        { status: "blocked", reason: "field-result" },
      ]);
      const manager = new FindReplaceManager<(typeof matches)[number]>();
      manager.setMatches(matches);
      for (const [index, foundMatch] of matches.entries()) {
        expect(manager.goTo(index)?.match).toEqual(foundMatch);
        expect(manager.replaceCurrent(source, "replaced")).toBeNull();
      }
      expect(manager.replaceAll(source, matches, "replaced")).toEqual({
        document: source,
        replacedCount: 0,
        skippedCount: 2,
      });
      expect(manager.getResult()?.matches).toEqual(matches);
      expect(source).toEqual(before);
      expect(toProseDoc(source).toJSON()).toEqual(beforePm);
    },
  );

  test.each(["simple", "complex"] as const)(
    "%s field hits are skipped while an ordinary body hit is replaced",
    (kind) => {
      const field = makeField(kind);
      const source: Document = {
        package: {
          document: { content: [{ type: "paragraph", content: [field, textRun(" hit")] }] },
        },
      };
      const matches = findInDocument(source, "it", createDefaultFindOptions());
      expect(matches).toHaveLength(3);
      const manager = new FindReplaceManager<(typeof matches)[number]>();
      manager.setMatches(matches);
      const result = manager.replaceAll(source, matches, "XX");
      if (!result) panic("Replacement outcome missing");
      expect(result).toMatchObject({ replacedCount: 1, skippedCount: 2 });
      expect(documentParagraph(result.document).content.at(0)).toEqual(field);
      expect(findInDocument(result.document, "hit hit", createDefaultFindOptions())).toHaveLength(
        1,
      );
      expect(findInDocument(result.document, "hXX", createDefaultFindOptions())).toHaveLength(1);
      const beforeFieldPm = toProseDoc(source).child(0).child(0).toJSON();
      expect(toProseDoc(result.document).child(0).child(0).toJSON()).toEqual(beforeFieldPm);
    },
  );

  test("a final-view match across hidden deletion text cannot destroy revision metadata", () => {
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
                  info: { id: 17, author: "Author", date: "2026-01-01T00:00:00Z" },
                  content: [textRun("X")],
                },
                textRun("B"),
              ],
            },
          ],
        },
      },
    };
    const before = structuredClone(source);
    const beforePm = toProseDoc(source).toJSON();
    const matches = findInDocument(source, "AB", {
      ...createDefaultFindOptions(),
      displayMode: "no-markup",
    });
    expect(matches).toHaveLength(1);
    expect(matches.at(0)).toMatchObject({ startOffset: 0, endOffset: 3, text: "AB" });
    expect(matches.at(0)?.replacement).toEqual({ status: "blocked", reason: "hidden-content" });
    const manager = new FindReplaceManager<(typeof matches)[number]>();
    manager.setMatches(matches);
    expect(manager.replaceCurrent(source, "replacement")).toBeNull();
    expect(manager.replaceAll(source, matches, "replacement")).toEqual({
      document: source,
      replacedCount: 0,
      skippedCount: 1,
    });
    expect(source).toEqual(before);
    expect(toProseDoc(source).toJSON()).toEqual(beforePm);
  });

  test("ordinary body matches remain replaceable individually and together", () => {
    const source: Document = {
      package: { document: { content: [{ type: "paragraph", content: [textRun("hit hit")] }] } },
    };
    const matches = findInDocument(source, "hit", createDefaultFindOptions());
    expect(matches.map(({ replacement }) => replacement)).toEqual([
      { status: "allowed" },
      { status: "allowed" },
    ]);
    const manager = new FindReplaceManager<(typeof matches)[number]>();
    manager.setMatches(matches);
    const current = manager.replaceCurrent(source, "value");
    if (!current) panic("Current body replacement failed");
    expect(findInDocument(current, "value hit", createDefaultFindOptions())).toHaveLength(1);
    const all = manager.replaceAll(source, matches, "value");
    if (!all) panic("Body replacement outcome missing");
    expect(all).toMatchObject({ replacedCount: 2, skippedCount: 0 });
    expect(findInDocument(all.document, "value value", createDefaultFindOptions())).toHaveLength(1);
    expect(findInDocument(source, "hit hit", createDefaultFindOptions())).toHaveLength(1);
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

    expect(manager.setMatches(matches)).toEqual({ matches, totalCount: 2, currentIndex: 0 });

    expect(manager.setMatches([])).toEqual({ matches: [], totalCount: 0, currentIndex: 0 });
  });

  test("navigate wraps the cursor and reports the now-current match", () => {
    const manager = new FindReplaceManager<TestMatch>();
    const matches = [match("a", 0, 0), match("b", 1, 0)];
    manager.setMatches(matches);

    expect(manager.navigate("next")).toEqual({ match: matches[1]!, index: 1 });
    expect(manager.navigate("next")).toEqual({ match: matches[0]!, index: 0 });
    expect(manager.navigate("previous")).toEqual({ match: matches[1]!, index: 1 });
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
    expect(result.currentIndex).toBe(Math.max(0, Math.min(index, matches.length - 1)));
    expect(manager.navigate("next")?.index).toBe((result.currentIndex + 1) % matches.length);
  }
  expect(manager.setMatches([], 4).currentIndex).toBe(0);
  expect(manager.navigate("next")).toBeNull();
});
