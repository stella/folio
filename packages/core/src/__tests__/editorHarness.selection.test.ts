import { expect, test } from "bun:test";
import { NodeSelection } from "prosemirror-state";

import { documentShape, shapeArrayBuffer } from "./documentShapes";
import { createHarnessState, parseShapeDocument, placeSelection } from "./editorHarness";

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
