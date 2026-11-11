import { expect } from "@playwright/test";

import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import {
  clearRefusals,
  expectNoRefusals,
  loadReady,
  reactPort,
  select,
  vuePort,
} from "../parity/canonicalInteractionAssertions";
import { canonicalLoadFixture } from "../parity/canonicalLoadFixture";
import { test } from "../parity/canonicalTest";

const cell = (text: string) => ({
  type: "tableCell" as const,
  content: [
    {
      type: "paragraph" as const,
      content: [
        { type: "run" as const, content: [{ type: "text" as const, text }] },
      ],
    },
  ],
});

const save = async (page: Parameters<typeof select>[0]) => {
  const saved = await page.evaluate(async () => {
    const buffer = await globalThis.__folioCanonical?.save();
    return buffer ? [...new Uint8Array(buffer)] : null;
  });
  if (!saved) throw new TypeError("The canonical document did not save.");
  return parseDocx(new Uint8Array(saved), {
    preloadFonts: false,
    detectVariables: false,
  });
};

const firstTable = (document: Awaited<ReturnType<typeof parseDocx>>) => {
  const table = document.package.document.content.find(
    (block) => block.type === "table",
  );
  if (table?.type !== "table") throw new TypeError("The table is absent.");
  return table;
};

test("table cell, row and table formatting from the editor's commands saves, and tracks in suggesting mode", async ({
  page,
}) => {
  test.setTimeout(90_000);
  const document = createEmptyDocument({ initialText: "Before" });
  document.package.document.content.push({
    type: "table",
    rows: [
      { type: "tableRow", cells: [cell("A"), cell("B")] },
      { type: "tableRow", cells: [cell("C"), cell("D")] },
    ],
  });
  const source = await canonicalLoadFixture(await createDocx(document));
  for (const port of [reactPort, vuePort]) {
    for (const suggesting of [false, true]) {
      await page.goto(`http://localhost:${port}/?session=canonical`);
      await page.waitForSelector(".layout-page");
      await clearRefusals(page);
      await loadReady(page, source);
      if (suggesting)
        expect(
          await page.evaluate(() =>
            globalThis.__folioCanonical?.setMode("suggesting"),
          ),
        ).toBe(true);
      // "Before" spans 0-8; the table opens at 8, its row at 9, the first cell
      // at 10 and its paragraph at 11, so "A" ends at 13.
      await select(page, 13);
      for (const [name, args] of [
        ["setCellFillColor", ["FFCC00"]],
        ["setCellVerticalAlign", ["center"]],
        ["setRowHeight", [600, "exact"]],
        ["setTableProperties", [{ justification: "center" }]],
      ] as const) {
        const ran = await page.evaluate(
          ({ name: command, args: commandArgs }) =>
            globalThis.__folioCanonical?.executeNamedCommand(command, [
              ...commandArgs,
            ]),
          { name, args },
        );
        expect(ran).toBe(true);
      }
      await expectNoRefusals(page);
      const table = firstTable(await save(page));
      const first = table.rows.at(0);
      if (suggesting) {
        // The suggested formatting is recorded as a change a reviewer can reject.
        expect(first?.cells.at(0)?.propertyChanges?.length).toBeGreaterThan(0);
        expect(first?.propertyChanges?.length).toBeGreaterThan(0);
        expect(table.propertyChanges?.length).toBeGreaterThan(0);
      }
      expect(first?.cells.at(0)?.formatting?.shading?.fill?.rgb).toBe("FFCC00");
      expect(first?.cells.at(0)?.formatting?.verticalAlign).toBe("center");
      expect(first?.formatting?.height?.value).toBe(600);
      expect(first?.formatting?.heightRule).toBe("exact");
      expect(table.formatting?.justification).toBe("center");
      // Other cells keep what they had.
      expect(first?.cells.at(1)?.formatting?.shading).toBeUndefined();
    }
  }
});
