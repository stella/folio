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

const textRange = (state: EditorState, text: string): { from: number; to: number } => {
  let range: { from: number; to: number } | null = null;
  state.doc.descendants((node, position) => {
    if (!range && node.isText && node.text === text) {
      range = { from: position, to: position + node.nodeSize };
    }
  });
  if (!range) {
    throw new Error(`Expected text node: ${text}`);
  }
  return range;
};

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

    expect(captured.type).toBe("capturedTextFormatting");
    expect(markNames(captured.marks)).toEqual(["bold", "fontFamily", "fontSize"]);
    expect(findMark(captured.marks, "fontFamily")?.attrs["ascii"]).toBe("Georgia");
    expect(findMark(captured.marks, "fontSize")?.attrs["size"]).toBe(24);
  });

  test("represents an effectively plain selection as a valid capture", () => {
    const plain = EditorState.create({
      schema,
      doc: schema.node("doc", null, [schema.node("paragraph", null, [schema.text("plain")])]),
    });
    expect(captureFormatMarks(select(plain, 1, 6))).toEqual({
      effectiveFormatting: {},
      marks: [],
      type: "capturedTextFormatting",
    });
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
    expect(captureFormatMarks(select(state, 1, 2))).toEqual({
      effectiveFormatting: {},
      marks: [],
      type: "capturedTextFormatting",
    });
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
    const override = findMark(captured.marks, "runFormattingOverride");
    expect(override?.attrs["bold"]).toBe(false);
    expect(override?.attrs["rtl"]).toBeNull();
  });

  test("does not synthesize a complex-script override through an inline content control", () => {
    const characterStyle = mark("characterStyle", { styleId: "ComplexToggle" });
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", { defaultTextFormatting: { boldCs: true } }, [
          schema.node("sdt", { sdtType: "richText" }, [schema.text("source", [characterStyle])]),
        ]),
      ]),
    });

    const captured = captureFormatMarks(select(state, 2, 8));

    expect(findMark(captured.marks, "characterStyle")?.attrs["styleId"]).toBe("ComplexToggle");
    expect(findMark(captured.marks, "runFormattingOverride")).toBeUndefined();
    expect(captured.effectiveFormatting.boldCs).toBeUndefined();
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
    expect(markNames(target)).toEqual([
      "bold",
      "fontFamily",
      "fontSize",
      "runFormattingOverride",
    ]);
    expect(findMark(target, "fontFamily")?.attrs["ascii"]).toBe("Georgia");
    expect(findMark(target, "fontSize")?.attrs["size"]).toBe(24);
    expect(findMark(target, "runFormattingOverride")?.attrs).toMatchObject({
      _authoredOn: ["bold"],
      _authoredValues: {
        fontFamily: { ascii: "Georgia", hAnsi: "Georgia" },
        fontSize: 24,
      },
    });
  });

  test("paints an effectively plain capture and materializes every safe target cancellation", () => {
    const targetDefaults = {
      allCaps: true,
      bold: true,
      boldCs: true,
      color: { rgb: "336699" },
      cs: true,
      doubleStrike: true,
      effect: "shimmer" as const,
      emboss: true,
      emphasisMark: "dot" as const,
      highlight: "yellow" as const,
      imprint: true,
      italic: true,
      italicCs: true,
      kerning: 8,
      outline: true,
      position: 4,
      scale: 120,
      shading: { fill: { rgb: "00AA00" } },
      shadow: true,
      smallCaps: true,
      spacing: 20,
      strike: true,
      underline: { style: "single" as const },
      vertAlign: "superscript" as const,
    };
    let state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("source")]),
        schema.node("paragraph", { defaultTextFormatting: targetDefaults }, [
          schema.text("target", [
            mark("bold"),
            mark("italic"),
            mark("underline", { style: "single" }),
            mark("strike"),
            mark("textColor", { rgb: "336699" }),
            mark("highlight", { color: "yellow" }),
            mark("superscript"),
            mark("allCaps"),
            mark("smallCaps"),
            mark("characterSpacing", { spacing: 20, position: 4, scale: 120, kerning: 8 }),
            mark("runShading", { rgb: "00AA00" }),
            mark("emboss"),
            mark("imprint"),
            mark("textShadow"),
            mark("emphasisMark", { type: "dot" }),
            mark("textOutline"),
            mark("textEffect", { effect: "shimmer" }),
          ]),
        ]),
      ]),
    });
    const source = textRange(state, "source");
    const captured = captureFormatMarks(select(state, source.from, source.to));
    const capturedBefore = {
      effectiveFormatting: structuredClone(captured.effectiveFormatting),
      marks: captured.marks.map((capturedMark) => capturedMark.toJSON()),
    };
    const target = textRange(state, "target");
    state = select(state, target.from, target.to);

    const handled = applyFormatMarks(captured)(state, (tr) => {
      state = state.apply(tr);
    });

    expect(handled).toBe(true);
    const painted = marksInRange(state, target.from, target.to);
    expect(markNames(painted)).toEqual(["characterSpacing", "runFormattingOverride"]);
    expect(findMark(painted, "runFormattingOverride")?.attrs).toMatchObject({
      allCaps: false,
      bold: false,
      boldCs: false,
      color: "auto",
      cs: false,
      doubleStrike: false,
      effect: "none",
      emboss: false,
      emphasisMark: "none",
      highlight: "none",
      imprint: false,
      italic: false,
      italicCs: false,
      kerning: 0,
      outline: false,
      position: 0,
      scale: 100,
      shading: { pattern: "nil" },
      shadow: false,
      smallCaps: false,
      spacing: 0,
      strike: false,
      underline: "none",
      vertAlign: "baseline",
    });
    expect({
      effectiveFormatting: captured.effectiveFormatting,
      marks: captured.marks.map((capturedMark) => capturedMark.toJSON()),
    }).toEqual(capturedBefore);
  });

  test("materializes a separate target-relative delta for each selected paragraph", () => {
    let state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", { defaultTextFormatting: { bold: true } }, [
          schema.text("source", [mark("bold")]),
        ]),
        schema.node("paragraph", { defaultTextFormatting: { bold: true } }, [
          schema.text("first", [mark("bold")]),
        ]),
        schema.node("paragraph", { defaultTextFormatting: { underline: { style: "single" } } }, [
          schema.text("second", [mark("underline", { style: "single" })]),
        ]),
      ]),
    });
    const source = textRange(state, "source");
    const captured = captureFormatMarks(select(state, source.from, source.to));
    const first = textRange(state, "first");
    const second = textRange(state, "second");
    state = select(state, first.from, second.to);

    applyFormatMarks(captured)(state, (tr) => {
      state = state.apply(tr);
    });

    const firstMarks = marksInRange(state, first.from, first.to);
    const secondMarks = marksInRange(state, second.from, second.to);
    expect(findMark(firstMarks, "runFormattingOverride")).toBeUndefined();
    expect(findMark(firstMarks, "bold")).toBeDefined();
    expect(findMark(secondMarks, "bold")).toBeDefined();
    expect(findMark(secondMarks, "runFormattingOverride")?.attrs["underline"]).toBe("none");
  });

  test("refuses atomically when an absent source font cannot cancel the target context", () => {
    let state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("source")]),
        schema.node("paragraph", { defaultTextFormatting: { fontFamily: { ascii: "Arial" } } }, [
          schema.text("target", [mark("fontFamily", { ascii: "Arial" })]),
        ]),
      ]),
    });
    const source = textRange(state, "source");
    const captured = captureFormatMarks(select(state, source.from, source.to));
    const target = textRange(state, "target");
    state = select(state, target.from, target.to);
    const before = state.doc.toJSON();

    const handled = applyFormatMarks(captured)(state, (tr) => {
      state = state.apply(tr);
    });

    expect(handled).toBe(false);
    expect(state.doc.toJSON()).toEqual(before);
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

  test("no capture is a no-op and does not clear the target", () => {
    let state = buildState();
    state = select(state, 8, 13);
    const before = state.doc.toJSON();

    const handled = applyFormatMarks(null)(state, (tr) => {
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
