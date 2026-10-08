import type { ParagraphContent } from "../src/model/document";

type RangeStart = Extract<ParagraphContent["type"], "bookmarkStart" | `${string}RangeStart`>;
type RangeEnd<Start extends RangeStart> = Start extends "bookmarkStart"
  ? "bookmarkEnd"
  : Start extends `${infer Prefix}RangeStart`
    ? `${Prefix}RangeEnd`
    : never;
type RangeFixtureFactories = {
  [Start in RangeStart]: () => readonly [
    Extract<ParagraphContent, { type: Start }>,
    Extract<ParagraphContent, { type: RangeEnd<Start> }>,
  ];
};

/** Every source range class supplies its matching endpoints and authored metadata. */
export const RANGE_ANCHOR_FIXTURE_FACTORIES = {
  bookmarkStart: () => [
    { type: "bookmarkStart", id: 11, name: "target", displacedByCustomXml: "next" },
    { type: "bookmarkEnd", id: 11, displacedByCustomXml: "prev" },
  ],
  commentRangeStart: () => [
    { type: "commentRangeStart", id: 14 },
    { type: "commentRangeEnd", id: 14 },
  ],
  moveFromRangeStart: () => [
    {
      type: "moveFromRangeStart",
      id: 12,
      name: "source",
      author: "Reviewer",
      date: "2026-01-01T00:00:00Z",
    },
    { type: "moveFromRangeEnd", id: 12 },
  ],
  moveToRangeStart: () => [
    {
      type: "moveToRangeStart",
      id: 13,
      name: "destination",
      author: "Reviewer",
      date: "2026-01-01T00:00:00Z",
    },
    { type: "moveToRangeEnd", id: 13 },
  ],
} satisfies RangeFixtureFactories;
