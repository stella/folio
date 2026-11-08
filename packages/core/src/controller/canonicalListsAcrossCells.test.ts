/**
 * A list command over a selection that reaches into table cells applies to
 * every selected paragraph, in the body and in each cell, as the editor's
 * own command does.
 */
import { expect, test } from "bun:test";

import { createCanonicalEditorHarness } from "../../../../test/canonicalEditorHarness";
import type { BlockContent, Document, Paragraph } from "../types/document";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const source = (): Document => ({
  package: {
    document: {
      content: [
        paragraph("10000000", "Before"),
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [
                { type: "tableCell", content: [paragraph("21000000", "A")] },
                { type: "tableCell", content: [paragraph("22000000", "B")] },
              ],
            },
          ],
        },
        paragraph("30000000", "After"),
      ],
    },
  },
});

const paragraphsOf = (blocks: readonly BlockContent[]): Paragraph[] =>
  blocks.flatMap((block) => {
    if (block.type === "paragraph") return [block];
    if (block.type === "table")
      return block.rows.flatMap((row) =>
        row.cells.flatMap((cell) => paragraphsOf(cell.content)),
      );
    return [];
  });

for (const mode of ["editing", "suggesting"] as const)
  for (const command of ["toggleBulletList", "toggleNumberedList"] as const)
    test(`${command} over a selection into table cells lists every paragraph in ${mode} mode`, () => {
      const driver = createCanonicalEditorHarness(source(), mode);
      try {
        const size = driver.state.doc.content.size;
        driver.history.setSelection(1, size - 1);
        expect(
          driver.execute(driver.commandManager.requireCommand(command)()),
        ).toBe(true);
        expect(driver.refusals).toEqual([]);
        const listed = paragraphsOf(
          driver.snapshot().package.document.content,
        ).map(({ formatting }) =>
          formatting?.numPr?.kind === "reference"
            ? formatting.numPr.numId
            : null,
        );
        expect(listed).toHaveLength(4);
        expect(listed.every((numId) => numId !== null)).toBe(true);
        // One list runs through the body and the cells.
        expect(new Set(listed).size).toBe(1);
      } finally {
        driver.dispose();
      }
    });
