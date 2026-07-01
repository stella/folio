import { describe, expect, test } from "bun:test";
import type { Mark } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

import { schema } from "../schema";
import { applyFormatMarks, captureFormatMarks } from "./formatPainter";

const mark = (name: string, attrs?: Record<string, unknown>): Mark => {
  const type = schema.marks[name];
  if (!type) {
    throw new Error(`Expected mark type in schema: ${name}`);
  }
  return type.create(attrs);
};

/** Marks of the first text node inside [from, to). */
const marksInRange = (state: EditorState, from: number, to: number): readonly Mark[] => {
  let found: readonly Mark[] | null = null;
  state.doc.nodesBetween(from, to, (node) => {
    if (found) {
      return false;
    }
    if (node.isText) {
      found = node.marks;
      return false;
    }
    return true;
  });
  return found ?? [];
};

const markNames = (marks: readonly Mark[]): string[] => marks.map((m) => m.type.name).toSorted();

const findMark = (marks: readonly Mark[], name: string): Mark | undefined =>
  marks.find((m) => m.type.name === name);

/**
 * doc: paragraph["Georgia"(bold, Georgia, 24) + "Arial"(Arial, 20)]
 * positions: "Georgia" = [1, 8), "Arial" = [8, 13)
 */
const buildState = (): EditorState =>
  EditorState.create({
    schema,
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("Georgia", [
          mark("bold"),
          mark("fontFamily", { ascii: "Georgia", hAnsi: "Georgia" }),
          mark("fontSize", { size: 24 }),
        ]),
        schema.text("Arial", [
          mark("fontFamily", { ascii: "Arial", hAnsi: "Arial" }),
          mark("fontSize", { size: 20 }),
        ]),
      ]),
    ]),
  });

const select = (state: EditorState, from: number, to: number): EditorState =>
  state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));

describe("captureFormatMarks", () => {
  test("captures multiple marks at once with their attrs", () => {
    const state = select(buildState(), 1, 8);
    const captured = captureFormatMarks(state);

    expect(markNames(captured)).toEqual(["bold", "fontFamily", "fontSize"]);
    expect(findMark(captured, "fontFamily")?.attrs["ascii"]).toBe("Georgia");
    expect(findMark(captured, "fontSize")?.attrs["size"]).toBe(24);
  });

  test("returns an empty array when the selection carries no direct formatting", () => {
    const plain = EditorState.create({
      schema,
      doc: schema.node("doc", null, [schema.node("paragraph", null, [schema.text("plain")])]),
    });
    expect(captureFormatMarks(select(plain, 1, 6))).toEqual([]);
  });

  test("does not paint an override mark that only carries excluded (hidden/rtl) attrs", () => {
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("x", [mark("runFormattingOverride", { hidden: false, rtl: false })]),
        ]),
      ]),
    });
    expect(captureFormatMarks(select(state, 1, 2))).toEqual([]);
  });

  test("strips excluded attrs but keeps the rest of a mixed override mark", () => {
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("x", [mark("runFormattingOverride", { bold: false, rtl: false })]),
        ]),
      ]),
    });
    const captured = captureFormatMarks(select(state, 1, 2));
    const override = findMark(captured, "runFormattingOverride");
    expect(override?.attrs["bold"]).toBe(false);
    expect(override?.attrs["rtl"]).toBeNull();
  });
});

describe("applyFormatMarks", () => {
  test("paints captured marks onto the target range", () => {
    let state = buildState();
    const captured = captureFormatMarks(select(state, 1, 8));

    state = select(state, 8, 13);
    const handled = applyFormatMarks(captured)(state, (tr) => {
      state = state.apply(tr);
    });

    expect(handled).toBe(true);
    const target = marksInRange(state, 8, 13);
    expect(markNames(target)).toEqual(["bold", "fontFamily", "fontSize"]);
    expect(findMark(target, "fontFamily")?.attrs["ascii"]).toBe("Georgia");
    expect(findMark(target, "fontSize")?.attrs["size"]).toBe(24);
  });

  test("replaces conflicting same-type marks instead of merging duplicates", () => {
    let state = buildState();
    const captured = captureFormatMarks(select(state, 1, 8));

    state = select(state, 8, 13);
    applyFormatMarks(captured)(state, (tr) => {
      state = state.apply(tr);
    });

    const fontFamilies = marksInRange(state, 8, 13).filter((m) => m.type.name === "fontFamily");
    expect(fontFamilies).toHaveLength(1);
    expect(fontFamilies[0]?.attrs["ascii"]).toBe("Georgia");
    const fontSizes = marksInRange(state, 8, 13).filter((m) => m.type.name === "fontSize");
    expect(fontSizes).toHaveLength(1);
    expect(fontSizes[0]?.attrs["size"]).toBe(24);
  });

  test("leaves structural marks (comments) untouched while painting", () => {
    let state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("src", [mark("bold")]),
          schema.text("dst", [mark("comment", { commentId: 7 })]),
        ]),
      ]),
    });
    const captured = captureFormatMarks(select(state, 1, 4));

    state = select(state, 4, 7);
    applyFormatMarks(captured)(state, (tr) => {
      state = state.apply(tr);
    });

    const target = marksInRange(state, 4, 7);
    expect(findMark(target, "bold")).toBeDefined();
    expect(findMark(target, "comment")?.attrs["commentId"]).toBe(7);
  });

  test("empty capture is a no-op and does not clear the target", () => {
    let state = buildState();
    state = select(state, 8, 13);
    const before = state.doc.toJSON();

    const handled = applyFormatMarks([])(state, (tr) => {
      state = state.apply(tr);
    });

    expect(handled).toBe(false);
    expect(state.doc.toJSON()).toEqual(before);
  });

  test("collapsed target selection is a no-op", () => {
    let state = buildState();
    const captured = captureFormatMarks(select(state, 1, 8));

    state = select(state, 10, 10);
    const before = state.doc.toJSON();
    const handled = applyFormatMarks(captured)(state, (tr) => {
      state = state.apply(tr);
    });

    expect(handled).toBe(false);
    expect(state.doc.toJSON()).toEqual(before);
  });
});
