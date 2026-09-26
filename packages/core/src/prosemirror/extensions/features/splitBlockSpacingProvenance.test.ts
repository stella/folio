import { describe, expect, test } from "bun:test";

import { documentShape } from "../../../__tests__/documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
  saveHarnessState,
} from "../../../__tests__/editorHarness";

describe("Enter keeps where the new paragraph's spacing comes from", () => {
  test.each([
    // Spacing from `w:docDefaults`.
    ["header-footer", "Body paragraph under a header."],
    // No styles part: spacing from the built-in default paragraph style.
    ["bare-package", "Second paragraph of a bare letter."],
  ] as const)(
    "a paragraph added with Enter saves no spacing it only inherits (%s)",
    async (shapeId, focus) => {
      const base = await parseShapeDocument(await documentShape(shapeId).build());
      const before = placeSelection(createHarnessState(base, "editing"), focus, "caret-end");
      if (!before) {
        throw new Error("the shape has no focus paragraph");
      }
      const view = new HeadlessEditorView(before);
      expect(view.pressKey("Enter")).toBe(true);

      const { model } = await saveHarnessState(view.state, base);
      const blocks = model.package.document.content;
      const index = blocks.findIndex((block) => JSON.stringify(block).includes(focus));
      const added = blocks[index + 1];
      expect(added?.type).toBe("paragraph");
      expect(added?.type === "paragraph" ? added.formatting : undefined).toBeUndefined();
    },
  );
});
