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
    "%s builds reproducibly, validates and holds its focus paragraph",
    async (_id, shape) => {
      const bytes = await shape.build();
      expect(await shape.build()).toBe(bytes);
      const document = await parseShapeDocument(bytes);
      assertValidFolioDocumentModel(document, `shape ${shape.id}`);
      const state = createHarnessState(document, "editing");
      expect(findTextblock(state.doc, shape.focus)).not.toBeNull();
    },
  );
});
