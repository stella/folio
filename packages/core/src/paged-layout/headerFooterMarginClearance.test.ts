import { describe, expect, test } from "bun:test";

import type { FlowBlock, HeaderFooterContent, PageMargins } from "../layout-engine/types";
import {
  bodyBlocksClearSectionHeaderFooter,
  bodyMarginsClearHeaderFooter,
} from "./headerFooterMarginClearance";

const margins = (top: number, bottom: number): PageMargins => ({
  top,
  right: 72,
  bottom,
  left: 72,
  header: 24,
  footer: 24,
});

const story = (height: number, bodyTopClearance?: number): HeaderFooterContent => ({
  blocks: [],
  measures: [],
  height,
  visualTop: 0,
  visualBottom: height,
  marginPushTop: 0,
  marginPushBottom: height,
  ...(bodyTopClearance === undefined ? {} : { bodyTopClearance }),
  textSig: "",
});

describe("header and footer body margin clearance", () => {
  test.each([
    { height: 12, expected: 72 },
    { height: 48, expected: 72 },
    { height: 96, expected: 120 },
    { height: 180, expected: 204 },
  ])(
    "clears header heights from below through well above the top margin",
    ({ height, expected }) => {
      expect(
        bodyMarginsClearHeaderFooter({
          authoredMargins: margins(72, 72),
          preparedHeader: story(height),
          preparedFooter: undefined,
        }).top,
      ).toBe(expected);
    },
  );

  test.each([
    { height: 12, expected: 72 },
    { height: 48, expected: 72 },
    { height: 96, expected: 120 },
    { height: 180, expected: 204 },
  ])(
    "clears footer heights from below through well above the bottom margin",
    ({ height, expected }) => {
      expect(
        bodyMarginsClearHeaderFooter({
          authoredMargins: margins(72, 72),
          preparedHeader: undefined,
          preparedFooter: story(height),
        }).bottom,
      ).toBe(expected);
    },
  );

  test("honors positioned header clearance beyond its flow height", () => {
    expect(
      bodyMarginsClearHeaderFooter({
        authoredMargins: margins(72, 72),
        preparedHeader: story(12, 160),
        preparedFooter: undefined,
      }).top,
    ).toBe(160);
  });

  test("applies each section's inherited default stories to its authored margins", () => {
    const firstMargins = margins(72, 72);
    const secondMargins = margins(40, 50);
    const blocks: FlowBlock[] = [
      { kind: "sectionBreak", id: "s1", margins: firstMargins },
      { kind: "sectionBreak", id: "s2", margins: secondMargins },
    ];
    const headers = new Map([
      ["default-1", story(100)],
      ["default-2", story(140)],
      ["first", story(220)],
      ["even", story(240)],
    ]);

    const cleared = bodyBlocksClearSectionHeaderFooter(blocks, {
      authoredMargins: firstMargins,
      sectionHeaderFooterRefs: [
        { headerDefault: "default-1", headerFirst: "first", titlePg: true },
        { headerDefault: "default-2", headerEven: "even", evenAndOddHeaders: true },
      ],
      headerContentByRId: headers,
      footerContentByRId: undefined,
    });

    expect(cleared.map((block) => block.margins?.top)).toEqual([124, 164]);
    expect(blocks.map((block) => block.margins?.top)).toEqual([72, 40]);
  });
});
