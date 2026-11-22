import { expect, test } from "bun:test";

import { layoutDocument } from "../index";
import type { FlowBlock, ParagraphBlock, TableBlock } from "../types";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "./__tests__/fakeTextMeasure";
import { measureBlocks } from "./measureBlocks";
import {
  hasSettledRaisedTableWrap,
  settleRaisedTableWrap,
} from "./raisedTableWrap";

const paragraph = (id: string, lines: readonly string[]): ParagraphBlock => ({
  kind: "paragraph",
  id,
  attrs: { keepLines: true },
  runs: lines.flatMap((text, index) =>
    index === 0
      ? [{ kind: "text", text }]
      : [{ kind: "lineBreak" }, { kind: "text", text }],
  ),
});

const table = (topFromText: number): TableBlock => ({
  kind: "table",
  id: "floating-table",
  columnWidths: [60],
  rows: [{ cells: [{ blocks: [paragraph("cell", ["Cell"])] }] }],
  floating: { horzAnchor: "text", vertAnchor: "text", tblpY: 0, topFromText },
});

const layOut = (blocks: FlowBlock[], height: number) => {
  const options = {
    pageSize: { w: 300, h: height },
    margins: { top: 10, bottom: 10, left: 40, right: 40 },
  };
  const measure = (raisedTableWrap?: Parameters<typeof measureBlocks>[5]) =>
    measureBlocks(
      blocks,
      220,
      10,
      {
        pageWidth: 300,
        pageHeight: height,
        marginLeft: 40,
        marginRight: 40,
        marginBottom: 10,
      },
      undefined,
      raisedTableWrap,
    );
  return settleRaisedTableWrap({
    blocks,
    measures: measure(),
    layout: (measures) => layoutDocument(blocks, measures, options),
    measure: (raisedTableWrap) => measure({ raisedTableWrap }),
  });
};

test("wraps only preceding lines reached by the top distance of a floating table", () => {
  withFakeTextMeasure(
    () => {
      const result = layOut(
        [
          paragraph("lead", ["First", "Second"]),
          table(10),
          paragraph("anchor", ["Anchor"]),
        ],
        300,
      );
      const lead = result.measures.at(0);
      expect(
        lead?.kind === "paragraph" &&
          lead.lines.map((line) => line.leftOffset ?? 0),
      ).toEqual([0, 72]);
    },
    { charWidth: fixedCharWidth(5) },
  );
});

test("a raised table exclusion stops at a natural page break", () => {
  withFakeTextMeasure(
    () => {
      const result = layOut(
        [
          paragraph("first-page", ["First", "Second", "Third"]),
          paragraph("second-page", ["Fourth", "Fifth"]),
          table(50),
          paragraph("anchor", ["Anchor"]),
        ],
        60,
      );
      expect(
        result.layout.pages.at(0)?.fragments.map(({ blockId }) => blockId),
      ).toEqual(["first-page"]);
      const first = result.measures.at(0);
      const second = result.measures.at(1);
      expect(
        first?.kind === "paragraph" &&
          first.lines.map((line) => line.leftOffset ?? 0),
      ).toEqual([0, 0, 0]);
      expect(
        second?.kind === "paragraph" &&
          second.lines.map((line) => line.leftOffset ?? 0),
      ).toEqual([72, 72]);
    },
    { charWidth: fixedCharWidth(5) },
  );
});

test("an empty paragraph that ignores table exclusions keeps its ordinary measure", () => {
  withFakeTextMeasure(() => {
    const result = layOut(
      [paragraph("empty", []), table(50), paragraph("anchor", ["Anchor"])],
      300,
    );
    const empty = result.measures.at(0);
    expect(
      empty?.kind === "paragraph" && hasSettledRaisedTableWrap(empty),
    ).toBe(false);
  });
});
