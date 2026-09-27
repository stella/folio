import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { schema } from "./schema";
import { projectMarkupView } from "./markupViewProjection";
import { withDocumentNumbering } from "./plugins/documentNumbering";
import { withDocumentStyles } from "./plugins/documentStyles";

describe("markup view projection cache", () => {
  test("reconfiguring styles or numbering invalidates a same-document projection", () => {
    const insertion = schema.marks["insertion"]?.create({
      revisionId: 1,
      author: "Reviewer",
      date: "2026-01-01T00:00:00Z",
    });
    if (!insertion) {
      throw new Error("Expected insertion mark in schema");
    }
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("New", [insertion]), schema.text(" text")]),
    ]);
    const state = EditorState.create({ doc });
    const initial = projectMarkupView(state, "no-markup");
    expect(initial.type).toBe("resolved");
    expect(projectMarkupView(state, "no-markup")).toBe(initial);

    const styled = withDocumentStyles(state, {
      styles: [{ styleId: "Normal", type: "paragraph", default: true }],
    });
    const numbered = withDocumentNumbering(state, { abstractNums: [], nums: [] });
    for (const refreshed of [styled, numbered]) {
      expect(refreshed.doc).toBe(state.doc);
      const projection = projectMarkupView(refreshed, "no-markup");
      expect(projection.type).toBe("resolved");
      expect(projection).not.toBe(initial);
      expect(projectMarkupView(refreshed, "no-markup")).toBe(projection);
    }
  });
});
