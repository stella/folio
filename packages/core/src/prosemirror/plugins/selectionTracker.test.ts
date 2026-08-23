import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import { TextSelection, EditorState } from "prosemirror-state";

import { schema } from "../schema";
import { extractSelectionSnapshot } from "../selectionState";
import { extractSelectionContext } from "./selectionTracker";

describe("extractSelectionContext", () => {
  test("reports marks found anywhere in a non-collapsed selection", () => {
    const bold = schema.marks.bold;
    const underline = schema.marks.underline;
    if (!bold || !underline) {
      panic("Expected bold and underline marks in schema");
    }

    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("plain"),
        schema.text("bold", [bold.create()]),
        schema.text("underlined", [underline.create({ style: "single" })]),
      ]),
    ]);
    const state = EditorState.create({
      doc,
      schema,
      selection: TextSelection.create(doc, 1, 18),
    });

    const context = extractSelectionContext(state);

    expect(context.hasSelection).toBe(true);
    expect(context.textFormatting.bold).toBe(true);
    expect(context.textFormatting.underline).toMatchObject({ style: "single" });
  });

  test("prefers visible underline when hidden underline occurs first", () => {
    const underline = schema.marks.underline;
    if (!underline) {
      panic("Expected underline mark in schema");
    }
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("hidden", [underline.create({ style: "none" })]),
        schema.text("visible", [underline.create({ style: "single" })]),
      ]),
    ]);
    const state = EditorState.create({
      doc,
      schema,
      selection: TextSelection.create(doc, 1, 14),
    });

    expect(extractSelectionContext(state).textFormatting.underline).toMatchObject({
      style: "single",
    });
  });

  test("matches the canonical snapshot for shared formatting and paragraph fields", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          indentLeft: 0,
          hangingIndent: false,
          styleId: "Heading1",
        },
        [schema.text("Shared")],
      ),
    ]);
    const state = EditorState.create({
      doc,
      schema,
      selection: TextSelection.create(doc, 1, 7),
    });

    const snapshot = extractSelectionSnapshot(state);
    const context = extractSelectionContext(state);

    expect(context.textFormatting).toEqual(snapshot.textFormatting);
    expect(context.paragraphFormatting).toMatchObject(snapshot.paragraphFormatting);
    expect(context.paragraphFormatting.styleId).toBe(snapshot.styleId);
    expect(context.startParagraphIndex).toBe(snapshot.startParagraphIndex);
    expect(context.endParagraphIndex).toBe(snapshot.endParagraphIndex);
  });
});
