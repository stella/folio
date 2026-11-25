import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import { fromMarkdown } from "../../markdown/fromMarkdown";
import type {
  Document,
  ShadingProperties,
  Table,
  TableCellFormatting,
} from "../../types/document";

const SHADINGS = {
  white: { pattern: "clear", color: { auto: true }, fill: { rgb: "FFFFFF" } },
  automatic: { pattern: "clear", color: { auto: true }, fill: { auto: true } },
  red: { pattern: "clear", color: { auto: true }, fill: { rgb: "FF0000" } },
} as const satisfies Record<string, ShadingProperties>;

const cellOf = (document: Document) => {
  const table = document.package.document.content.find(
    (block) => block.type === "table",
  );
  const cell = table?.rows.at(0)?.cells.at(0);
  if (!cell) throw new TypeError("The producer fixture has its table cell");
  return cell;
};

const fixture = (formatting: TableCellFormatting) => {
  const document = fromMarkdown("Cell shading.");
  const table = {
    type: "table",
    columnWidths: [2000],
    rows: [
      {
        type: "tableRow",
        cells: [
          {
            type: "tableCell",
            formatting,
            content: [
              {
                type: "paragraph",
                paraId: "00000002",
                content: [
                  { type: "run", content: [{ type: "text", text: "cell" }] },
                ],
              },
            ],
          },
        ],
      },
    ],
  } satisfies Table;
  document.package.document.content = [table];
  return document;
};

for (const [name, shading] of Object.entries(SHADINGS))
  test(`rejecting ${name} cell shading restores the authored model through parse and save`, async () => {
    const previous = {
      width: { value: 2000, type: "dxa" },
      shading,
    } as const satisfies TableCellFormatting;
    const originalBytes = await createDocx(fixture(previous));
    const original = await parseDocx(originalBytes);
    const changed = fixture({
      width: { value: 2000, type: "dxa" },
      shading: {
        pattern: "clear",
        color: { auto: true },
        fill: { rgb: "0000FF" },
      },
    });
    cellOf(changed).propertyChanges = [
      {
        type: "tableCellPropertyChange",
        info: { id: 7, author: "Writer", date: "2026-01-02T03:04:05Z" },
        previousFormatting: previous,
      },
    ];
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(changed),
    );
    const baseline = await FolioDocxReviewer.fromBuffer(originalBytes);
    expect(reviewer.rejectAll()).toEqual({ resolved: 1, skipped: [] });
    const actual = reviewer
      .storyTables({ view: "current-markup" })
      .at(0)
      ?.node.child(0)
      .child(0);
    const expected = baseline
      .storyTables({ view: "current-markup" })
      .at(0)
      ?.node.child(0)
      .child(0);
    expect(actual).toBeDefined();
    expect(actual?.attrs["backgroundColor"]).toBe(
      expected?.attrs["backgroundColor"],
    );
    expect(actual?.attrs["_resolvedBackgroundColor"]).toBe(
      expected?.attrs["_resolvedBackgroundColor"],
    );
    expect(cellOf(reviewer.toDocument()).formatting).toStrictEqual(
      cellOf(original).formatting,
    );
    const reopened = await parseDocx(await reviewer.toBuffer());
    expect(cellOf(reopened).formatting).toStrictEqual(
      cellOf(original).formatting,
    );
    expect(cellOf(reopened).propertyChanges).toBeUndefined();
  });
