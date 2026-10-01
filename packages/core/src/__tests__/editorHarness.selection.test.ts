import { expect, test } from "bun:test";
import { NodeSelection } from "prosemirror-state";

import { DOCUMENT_SHAPES, documentShape, shapeArrayBuffer } from "./documentShapes";
import {
  createHarnessState,
  parseShapeDocument,
  placeSelection,
  TEXTBLOCK_SELECTION_PLACEMENTS,
} from "./editorHarness";

test("node placement selects an inline atom and declines a text-only paragraph", async () => {
  for (const shapeId of ["image", "plain-markdown"]) {
    const shape = documentShape(shapeId);
    const model = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer(shape)));
    const state = createHarnessState(model, "editing");
    const placed = placeSelection(state, shape.focus, "node");
    if (shapeId === "plain-markdown") {
      expect(placed).toBeNull();
      continue;
    }
    expect(placed).not.toBeNull();
    expect(placed?.selection).toBeInstanceOf(NodeSelection);
    if (placed?.selection instanceof NodeSelection) {
      expect(placed.selection.node.isAtom).toBe(true);
      expect(placed.selection.node.isText).toBe(false);
    }
  }
});

// Text-only paste matrices require every declared placement to materialize;
// atom placement has its separate image fixture above.
test.each(DOCUMENT_SHAPES)("textblock placements materialize for $id", async (shape) => {
  const model = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer(shape)));
  const state = createHarnessState(model, "editing");
  for (const placement of TEXTBLOCK_SELECTION_PLACEMENTS) {
    expect(placeSelection(state, shape.focus, placement)).not.toBeNull();
  }
});
