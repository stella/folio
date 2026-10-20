import { describe, expect, test } from "bun:test";
import { CellSelection } from "prosemirror-tables";

import { createHarnessState, parseShapeDocument } from "../__tests__/editorHarness";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { canonicalSelectionRanges } from "./canonicalSelectionRange";

const stateFor = async (markdown: string) =>
  createHarnessState(
    await parseShapeDocument(new Uint8Array(await createDocx(fromMarkdown(markdown)))),
    "editing",
  );

const cellPositions = (doc: Parameters<typeof CellSelection.create>[0]) => {
  const positions = new Map<string, number>();
  doc.descendants((node, position) => {
    if (node.type.name === "tableCell") positions.set(node.textContent, position);
  });
  return positions;
};

describe("canonical cell selection ranges", () => {
  test.each([
    {
      name: "B1:B2 in a 2x2 table",
      markdown: "| H1 | H2 |\n|---|---|\n| A1 | B1 |\n| A2 | B2 |",
      anchor: "B1",
      head: "B2",
      expected: ["B1", "B2"],
    },
    {
      name: "B1:C3 in a 3x3 table",
      markdown:
        "| H1 | H2 | H3 |\n|---|---|---|\n| A1 | B1 | C1 |\n| A2 | B2 | C2 |\n| A3 | B3 | C3 |",
      anchor: "B1",
      head: "C3",
      expected: ["B1", "C1", "B2", "C2", "B3", "C3"],
    },
  ])("returns only $name cells in document order", async ({ markdown, anchor, head, expected }) => {
    const state = await stateFor(markdown);
    const positions = cellPositions(state.doc);
    const selection = CellSelection.create(
      state.doc,
      positions.get(anchor) ?? -1,
      positions.get(head) ?? -1,
    );
    const selected = canonicalSelectionRanges(state.apply(state.tr.setSelection(selection))).map(
      ({ from, to }) => state.doc.textBetween(from, to),
    );

    expect(selected).toEqual(expected);
  });
});
