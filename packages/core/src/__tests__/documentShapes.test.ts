import { describe, expect, test } from "bun:test";

import { assertValidFolioDocumentModel } from "../docx/modelValidation";
import { DOCUMENT_SHAPES } from "./documentShapes";
import { createHarnessState, findTextblock, parseShapeDocument } from "./editorHarness";

describe("document shapes", () => {
  test("ids are unique", () => {
    const ids = DOCUMENT_SHAPES.map((shape) => shape.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  test.each(DOCUMENT_SHAPES.map((shape) => [shape.id, shape] as const))(
    "%s builds the same content twice, validates and holds its focus paragraph",
    async (_id, shape) => {
      const document = await parseShapeDocument(await shape.build());
      const again = await parseShapeDocument(await shape.rebuild());
      expect(JSON.stringify(again.package.document)).toBe(
        JSON.stringify(document.package.document),
      );
      assertValidFolioDocumentModel(document, `shape ${shape.id}`);
      const state = createHarnessState(document, "editing");
      expect(findTextblock(state.doc, shape.focus)).not.toBeNull();
    },
  );
});
