import { describe, expect, test } from "bun:test";
import { paragraphLogicalText } from "@stll/docx-core/ops";
import type { Document, Paragraph, TableCell } from "@stll/docx-core/model";
import { CellSelection } from "prosemirror-tables";

import { createCanonicalEditorHarness } from "../../../../test/canonicalEditorHarness";
import { clearFormatting } from "../prosemirror/commands/formatting";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { EditorMode } from "./editorHarness";

const TABLES = [
  {
    name: "B1:B2 in a 2x2 table",
    markdown: "| A1 | B1 |\n|---|---|\n| A2 | B2 |",
    anchor: "B1",
    head: "B2",
    selected: ["B1", "B2"],
  },
  {
    name: "B1:C3 in a 3x3 table",
    markdown: "| A1 | B1 | C1 |\n|---|---|---|\n| A2 | B2 | C2 |\n| A3 | B3 | C3 |",
    anchor: "B1",
    head: "C3",
    selected: ["B1", "C1", "B2", "C2", "B3", "C3"],
  },
  {
    name: "A2:B3 in a 3x3 table",
    markdown: "| A1 | B1 | C1 |\n|---|---|---|\n| A2 | B2 | C2 |\n| A3 | B3 | C3 |",
    anchor: "A2",
    head: "B3",
    selected: ["A2", "B2", "A3", "B3"],
  },
] as const;

const MODES: readonly EditorMode[] = ["editing", "suggesting"];
const OPERATIONS = ["delete", "cut", "bold", "clear-formatting"] as const;

const tableCells = (document: Document): TableCell[] => {
  const table = document.package.document.content.at(0);
  if (table?.type !== "table") throw new TypeError("Expected the table fixture.");
  return table.rows.flatMap(({ cells }) => cells);
};

const cellParagraph = (cell: TableCell): Paragraph => {
  const paragraph = cell.content.at(0);
  if (paragraph?.type !== "paragraph") throw new TypeError("Expected one paragraph per cell.");
  return paragraph;
};

const prepareTable = (markdown: string): Document => {
  const document = fromMarkdown(markdown);
  for (const [index, cell] of tableCells(document).entries()) {
    const paragraph = cellParagraph(cell);
    paragraph.paraId = (index + 1).toString(16).padStart(8, "0").toUpperCase();
    paragraph.formatting = { alignment: "center" };
    for (const content of paragraph.content) {
      if (content.type === "run") {
        content.formatting = { italic: true, color: { rgb: "C00000" } };
      }
    }
  }
  return document;
};

const cellsByOriginalText = (document: Document) =>
  new Map(tableCells(document).map((cell) => [paragraphLogicalText(cellParagraph(cell)), cell]));

const cellPositions = (driver: ReturnType<typeof createCanonicalEditorHarness>) => {
  const positions = new Map<string, number>();
  driver.state.doc.descendants((node, position) => {
    if (node.type.name === "tableCell") positions.set(node.textContent, position);
  });
  return positions;
};

const proseCells = (driver: ReturnType<typeof createCanonicalEditorHarness>) => {
  const cells: (typeof driver.state.doc)[] = [];
  driver.state.doc.descendants((node) => {
    if (node.type.name === "tableCell") cells.push(node);
  });
  return cells;
};

describe("canonical cell-selection commands", () => {
  test.each(
    TABLES.flatMap((table) =>
      MODES.flatMap((mode) =>
        OPERATIONS.map(
          (operation) => [`${table.name}, ${operation}, ${mode}`, table, operation, mode] as const,
        ),
      ),
    ),
  )("%s", (_name, table, operation, mode) => {
    const driver = createCanonicalEditorHarness(prepareTable(table.markdown), mode);
    try {
      const positions = cellPositions(driver);
      const selection = CellSelection.create(
        driver.state.doc,
        positions.get(table.anchor) ?? -1,
        positions.get(table.head) ?? -1,
      );
      driver.dispatch(driver.state.tr.setSelection(selection));
      const beforeSelection = driver.state.selection.toJSON();
      const before = driver.snapshot();
      const beforeCells = cellsByOriginalText(before);

      if (operation === "delete") driver.pressKey("Delete");
      else if (operation === "cut") expect(driver.cut()).toBe(true);
      else if (operation === "bold")
        expect(driver.execute(driver.commandManager.requireCommand("toggleBold")())).toBe(true);
      else expect(driver.execute(clearFormatting)).toBe(true);

      const after = driver.snapshot();
      const afterCells = tableCells(after);
      const selected = new Set<string>(table.selected);
      for (const [index, afterCell] of afterCells.entries()) {
        const beforeCell = tableCells(before).at(index);
        if (!beforeCell) throw new TypeError("The command changed the table dimensions.");
        const label = paragraphLogicalText(cellParagraph(beforeCell));
        if (!selected.has(label)) {
          expect(afterCell).toStrictEqual(beforeCells.get(label));
          continue;
        }
        const beforeParagraph = cellParagraph(beforeCell);
        const afterParagraph = cellParagraph(afterCell);
        expect(afterParagraph.paraId).toBe(beforeParagraph.paraId);
        if ((operation === "delete" || operation === "cut") && mode === "editing") {
          expect(paragraphLogicalText(afterParagraph)).toBe("");
        } else if (operation === "clear-formatting") {
          expect(afterParagraph.formatting).toEqual(beforeParagraph.formatting);
          expect(
            afterParagraph.content.every(
              (content) =>
                content.type !== "run" || Object.keys(content.formatting ?? {}).length === 0,
            ),
          ).toBe(true);
        }
      }

      const projectedCells = proseCells(driver);
      for (const [index, cell] of projectedCells.entries()) {
        const sourceCell = tableCells(before).at(index);
        if (!sourceCell) throw new TypeError("The projection changed the table dimensions.");
        const label = paragraphLogicalText(cellParagraph(sourceCell));
        if (!selected.has(label)) continue;
        if ((operation === "delete" || operation === "cut") && mode === "suggesting") {
          expect(cell.textContent).toBe(label);
          cell.descendants((node) => {
            if (node.isText)
              expect(node.marks.some(({ type }) => type.name === "deletion")).toBe(true);
          });
        }
        if (operation === "bold") {
          cell.descendants((node) => {
            if (node.isText) expect(node.marks.some(({ type }) => type.name === "bold")).toBe(true);
          });
        }
      }

      const changed = driver.snapshot();
      const changedSelection = driver.state.selection.toJSON();
      expect(driver.history.undo()).toBe(true);
      expect(driver.snapshot()).toStrictEqual(before);
      expect(driver.state.selection.toJSON()).toEqual(beforeSelection);
      expect(driver.history.redo()).toBe(true);
      expect(driver.snapshot()).toStrictEqual(changed);
      expect(driver.state.selection.toJSON()).toEqual(changedSelection);
    } finally {
      driver.dispose();
    }
  });
});
