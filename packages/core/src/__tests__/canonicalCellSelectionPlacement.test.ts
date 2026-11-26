import { describe, expect, test } from "bun:test";
import { CellSelection } from "prosemirror-tables";

import type { Document, Paragraph, Table } from "@stll/docx-core/model";
import { paragraphLogicalText } from "@stll/docx-core/ops";

import { createCanonicalEditorHarness } from "../../../../test/canonicalEditorHarness";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { EditorMode } from "./editorHarness";

const MODES: readonly EditorMode[] = ["editing", "suggesting"];

const prepare = (): Document => {
  const document = fromMarkdown(
    "# Overview\n\n| A1 | B1 | C1 |\n|---|---|---|\n| A2 | B2 | C2 |\n| A3 | B3 | C3 |",
  );
  let next = 1;
  const identify = (paragraph: Paragraph) => {
    paragraph.paraId = (next++).toString(16).padStart(8, "0").toUpperCase();
  };
  for (const block of document.package.document.content) {
    if (block.type === "paragraph") identify(block);
    if (block.type === "table")
      for (const row of block.rows)
        for (const cell of row.cells)
          for (const paragraph of cell.content)
            if (paragraph.type === "paragraph") identify(paragraph);
  }
  return document;
};

const table = (document: Document): Table => {
  const found = document.package.document.content.find(
    (block) => block.type === "table",
  );
  if (found?.type !== "table") throw new TypeError("Expected the table.");
  return found;
};

/** Every cell's paragraphs by the text the cell held in the fixture. */
const cellParagraphs = (document: Document, labels: readonly string[]) =>
  new Map(
    table(document)
      .rows.flatMap(({ cells }) => cells)
      .map((cell, index) => [
        labels.at(index) ?? "",
        cell.content.filter(
          (block): block is Paragraph => block.type === "paragraph",
        ),
      ]),
  );

const LABELS = ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"];

const selectCells = (
  driver: ReturnType<typeof createCanonicalEditorHarness>,
  anchor: string,
  head: string,
) => {
  const positions = new Map<string, number>();
  driver.state.doc.descendants((node, position) => {
    if (node.type.name === "tableCell")
      positions.set(node.textContent, position);
  });
  driver.dispatch(
    driver.state.tr.setSelection(
      CellSelection.create(
        driver.state.doc,
        positions.get(anchor) ?? -1,
        positions.get(head) ?? -1,
      ),
    ),
  );
};

/** A paragraph's note references, inside a suggested insertion too. */
const noteReferenceCount = (paragraph: Paragraph) =>
  JSON.stringify(paragraph).match(/"type":"(?:footnoteRef|endnoteRef)"/g)
    ?.length ?? 0;

describe("notes and contents over selected cells", () => {
  test.each(
    MODES.flatMap((mode) =>
      (["insertFootnote", "insertEndnote"] as const).map(
        (command) => [`${command}, ${mode}`, command, mode] as const,
      ),
    ),
  )(
    "%s places the reference after the last selected cell's text",
    (_name, command, mode) => {
      const driver = createCanonicalEditorHarness(prepare(), mode);
      try {
        selectCells(driver, "B1", "C2");
        const before = driver.snapshot();
        expect(
          driver.execute(driver.commandManager.requireCommand(command)(1)),
        ).toBe(true);
        const after = driver.snapshot();
        const cells = cellParagraphs(after, LABELS);
        for (const label of LABELS)
          expect(
            (cells.get(label) ?? []).reduce(
              (count, paragraph) => count + noteReferenceCount(paragraph),
              0,
            ),
          ).toBe(label === "C2" ? 1 : 0);
        const last = cells.get("C2")?.at(0);
        if (last === undefined) throw new TypeError("Expected C2's paragraph.");
        // The reference follows the cell's text.
        const serialized = JSON.stringify(last);
        expect(serialized.indexOf('"C2"')).toBeGreaterThan(-1);
        expect(serialized.indexOf('"C2"')).toBeLessThan(
          serialized.search(/"type":"(?:footnoteRef|endnoteRef)"/),
        );
        expect(paragraphLogicalText(last).startsWith("C2")).toBe(true);
        const notes =
          command === "insertFootnote"
            ? after.package.footnotes
            : after.package.endnotes;
        expect(notes?.some((note) => note.id === 1)).toBe(true);
        expect(driver.history.undo()).toBe(true);
        expect(driver.snapshot()).toStrictEqual(before);
      } finally {
        driver.dispose();
      }
    },
  );

  test.each(MODES.map((mode) => [mode] as const))(
    "a table of contents goes at the start of the first selected cell, %s",
    (mode) => {
      const driver = createCanonicalEditorHarness(prepare(), mode);
      try {
        selectCells(driver, "C2", "B1");
        const before = driver.snapshot();
        expect(
          driver.execute(
            driver.commandManager.requireCommand("generateTOC")({
              title: "Contents",
            }),
          ),
        ).toBe(true);
        const after = driver.snapshot();
        const cells = cellParagraphs(after, LABELS);
        const first = cells.get("B1") ?? [];
        expect(first.length).toBeGreaterThan(1);
        expect(first.map(paragraphLogicalText)).toContain("Contents");
        // The cell's own text follows the contents.
        expect(first.map(paragraphLogicalText).at(-1)).toBe("B1");
        for (const label of LABELS)
          if (label !== "B1") expect(cells.get(label)).toHaveLength(1);
        expect(driver.history.undo()).toBe(true);
        expect(driver.snapshot()).toStrictEqual(before);
      } finally {
        driver.dispose();
      }
    },
  );
});
