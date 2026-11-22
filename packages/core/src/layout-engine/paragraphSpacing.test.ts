import { describe, expect, test } from "bun:test";

import { withFakeTextMeasure } from "./measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "./measure/measureBlocks";
import {
  collapseParagraphSpacing,
  resolveBodyTableLeadingSpacing,
  resolveEffectiveParagraphSpacing,
  resolveEffectiveParagraphSpacingTree,
} from "./paragraphSpacing";
import type {
  ParagraphAttrs,
  ParagraphBlock,
  TableBlock,
  TextBoxBlock,
} from "./types";

const paragraph = (
  id: string,
  attrs: ParagraphAttrs,
  text = id,
): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [{ kind: "text", text }],
  attrs,
});

describe("resolveEffectiveParagraphSpacing", () => {
  test("collapses only inherited spacing on an otherwise bare empty paragraph", () => {
    const inherited = paragraph(
      "inherited",
      { spacing: { before: 12, after: 18 } },
      "",
    );
    const explicit = paragraph(
      "explicit",
      {
        spacing: { before: 12, after: 18 },
        spacingExplicit: { before: true },
      },
      "",
    );

    expect(resolveEffectiveParagraphSpacing(inherited)).toEqual({
      before: 0,
      after: 0,
    });
    expect(resolveEffectiveParagraphSpacing(explicit)).toEqual({
      before: 12,
      after: 0,
    });
  });

  test("preserves inherited spacing when the empty paragraph has authored provenance", () => {
    expect(
      resolveEffectiveParagraphSpacing(
        paragraph(
          "styled",
          { styleId: "Body", spacing: { before: 12, after: 18 } },
          "",
        ),
      ),
    ).toEqual({ before: 12, after: 18 });
  });
});

describe("resolveEffectiveParagraphSpacingTree", () => {
  test("derives contextual spacing without mutating authored paragraphs", () => {
    const first = paragraph("first", {
      styleId: "Body",
      contextualSpacing: true,
      spacing: { before: 5, after: 13 },
    });
    const second = paragraph("second", {
      styleId: "Body",
      contextualSpacing: true,
      spacing: { before: 5, after: 13 },
    });

    const resolved = resolveEffectiveParagraphSpacingTree([first, second]);

    expect(first.attrs?.spacing).toEqual({ before: 5, after: 13 });
    expect(second.attrs?.spacing).toEqual({ before: 5, after: 13 });
    expect(resolved.map((block) => block.attrs?.spacing)).toEqual([
      { before: 5, after: 0 },
      { before: 0, after: 13 },
    ]);
    const repeated = resolveEffectiveParagraphSpacingTree(resolved);
    expect(repeated).toBe(resolved);
    expect(repeated.at(0)).toBe(resolved.at(0));
    expect(repeated.at(1)).toBe(resolved.at(1));
  });

  test("suppresses automatic spacing only inside one numbered sequence", () => {
    const first = paragraph("first", {
      automaticSpacing: { after: true },
      numPr: { kind: "reference", numId: 4, ilvl: 0 },
      spacing: { before: 18, after: 18 },
    });
    const second = paragraph("second", {
      automaticSpacing: { before: true },
      numPr: { kind: "reference", numId: 4, ilvl: 0 },
      spacing: { before: 18, after: 18 },
    });
    const boundary = paragraph("boundary", {
      automaticSpacing: { before: true },
      numPr: { kind: "reference", numId: 5, ilvl: 0 },
      spacing: { before: 18, after: 18 },
    });

    const resolved = resolveEffectiveParagraphSpacingTree([
      first,
      second,
      boundary,
    ]);

    expect(resolved.map((block) => block.attrs?.spacing)).toEqual([
      { before: 18, after: 0 },
      { before: 0, after: 18 },
      { before: 18, after: 18 },
    ]);
  });

  test.each([
    ["Cell", 0, 0],
    ["Other", 8, 6],
  ] as const)(
    "resolves contextual spacing at the edge of the adjacent %s cell",
    (rightStyle, leftAfter, rightBefore) => {
      const left = paragraph("left", {
        styleId: "Cell",
        contextualSpacing: true,
        spacing: { after: 8 },
      });
      const right = paragraph("right", {
        styleId: rightStyle,
        contextualSpacing: true,
        spacing: { before: 6 },
      });
      const table: TableBlock = {
        kind: "table",
        id: "table",
        rows: [
          {
            id: "row",
            cells: [
              { id: "left-cell", blocks: [left] },
              { id: "right-cell", blocks: [right] },
            ],
          },
        ],
      };

      const resolved = resolveEffectiveParagraphSpacingTree([table]).at(0);
      if (resolved?.kind !== "table") {
        throw new Error("Expected table block");
      }

      expect(
        resolved.rows.at(0)?.cells.map((cell) => {
          const block = cell.blocks.at(0);
          return block?.kind === "paragraph" ? block.attrs?.spacing : undefined;
        }),
      ).toEqual([{ after: leftAfter }, { before: rightBefore }]);
      expect(left.attrs?.spacing).toEqual({ after: 8 });
      expect(right.attrs?.spacing).toEqual({ before: 6 });
      expect(resolveEffectiveParagraphSpacingTree([resolved]).at(0)).toBe(
        resolved,
      );
    },
  );

  test("keeps spacing between cells separated by a vertical merge", () => {
    const left = paragraph("left", {
      styleId: "Cell",
      contextualSpacing: true,
      spacing: { after: 8 },
    });
    const right = paragraph("right", {
      styleId: "Cell",
      contextualSpacing: true,
      spacing: { before: 6 },
    });
    const table: TableBlock = {
      kind: "table",
      id: "table",
      columnWidths: [60, 60, 60],
      rows: [
        {
          id: "first-row",
          cells: [
            { id: "first-left", blocks: [] },
            { id: "merged", rowSpan: 2, blocks: [] },
            { id: "first-right", blocks: [] },
          ],
        },
        {
          id: "second-row",
          cells: [
            { id: "left-cell", blocks: [left] },
            { id: "right-cell", blocks: [right] },
          ],
        },
      ],
    };

    const resolved = resolveEffectiveParagraphSpacingTree([table]).at(0);

    expect(resolved).toBe(table);
    expect(left.attrs?.spacing?.after).toBe(8);
    expect(right.attrs?.spacing?.before).toBe(6);
  });

  test("resolves a table's leading contextual spacing against the current lead style", () => {
    withFakeTextMeasure(() => {
      const first = paragraph("first", {
        styleId: "A",
        contextualSpacing: true,
        spacing: { before: 6 },
      });
      const table: TableBlock = {
        kind: "table",
        id: "table",
        rows: [{ id: "row", cells: [{ id: "cell", blocks: [first] }] }],
      };

      const layoutWithLeadStyle = (styleId: string) => {
        const lead = paragraph("lead", {
          styleId,
          contextualSpacing: true,
          spacing: { after: 7 },
        });
        const blocks = resolveBodyTableLeadingSpacing([lead, table]);
        const [resolvedLead, resolvedTable] = blocks;
        if (
          resolvedLead?.kind !== "paragraph" ||
          resolvedTable?.kind !== "table"
        ) {
          throw new Error("Expected lead paragraph and table");
        }
        const cellStart = resolvedTable.rows.at(0)?.cells.at(0)?.blocks.at(0);
        const measure = measureBlocks(blocks, 220).at(1);
        if (cellStart?.kind !== "paragraph" || measure?.kind !== "table") {
          throw new Error("Expected a measured table with a cell paragraph");
        }
        expect(resolvedLead.attrs?.spacing?.after).toBe(7);
        return {
          before: cellStart.attrs?.spacing?.before,
          height: measure.totalHeight,
        };
      };

      const firstMatch = layoutWithLeadStyle("A");
      const different = layoutWithLeadStyle("B");
      const secondMatch = layoutWithLeadStyle("A");

      expect([firstMatch.before, different.before, secondMatch.before]).toEqual(
        [0, 6, 0],
      );
      expect(different.height - firstMatch.height).toBeCloseTo(6);
      expect(secondMatch.height).toBe(firstMatch.height);
      expect(first.attrs?.spacing?.before).toBe(6);
    });
  });

  test("resolves nested table and text-box stories with structural sharing", () => {
    const tableFirst = paragraph("table-first", {
      styleId: "Cell",
      contextualSpacing: true,
      spacing: { after: 8 },
    });
    const tableSecond = paragraph("table-second", {
      styleId: "Cell",
      contextualSpacing: true,
      spacing: { before: 6 },
    });
    const table: TableBlock = {
      kind: "table",
      id: "table",
      rows: [
        {
          id: "row",
          cells: [{ id: "cell", blocks: [tableFirst, tableSecond] }],
        },
      ],
    };
    const textBoxFirst = paragraph("box-first", {
      styleId: "Box",
      contextualSpacing: true,
      spacing: { after: 9 },
    });
    const textBoxSecond = paragraph("box-second", {
      styleId: "Box",
      contextualSpacing: true,
      spacing: { before: 7 },
    });
    const textBox: TextBoxBlock = {
      kind: "textBox",
      id: "box",
      width: 200,
      content: [textBoxFirst, textBoxSecond],
    };
    const unchanged = paragraph("unchanged", { spacing: { after: 4 } });

    const resolved = resolveEffectiveParagraphSpacingTree([
      table,
      textBox,
      unchanged,
    ]);

    expect(resolved.at(0)).not.toBe(table);
    expect(resolved.at(1)).not.toBe(textBox);
    expect(resolved.at(2)).toBe(unchanged);
    const resolvedTable = resolved.at(0);
    const resolvedTextBox = resolved.at(1);
    expect(resolvedTable?.kind).toBe("table");
    expect(resolvedTextBox?.kind).toBe("textBox");
    if (
      resolvedTable?.kind !== "table" ||
      resolvedTextBox?.kind !== "textBox"
    ) {
      return;
    }
    expect(
      resolvedTable.rows
        .at(0)
        ?.cells.at(0)
        ?.blocks.map((block) => block.attrs?.spacing),
    ).toEqual([{ after: 0 }, { before: 0 }]);
    expect(
      resolvedTextBox.content.map((block) =>
        block.kind === "paragraph" ? block.attrs?.spacing : undefined,
      ),
    ).toEqual([{ after: 0 }, { before: 0 }]);
    expect(tableFirst.attrs?.spacing?.after).toBe(8);
    expect(textBoxFirst.attrs?.spacing?.after).toBe(9);
  });
});

describe("collapseParagraphSpacing", () => {
  test("uses the larger side of a paragraph boundary", () => {
    expect(collapseParagraphSpacing({ before: 7, after: 11 })).toBe(11);
  });
});
