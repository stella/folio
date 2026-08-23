import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Command } from "prosemirror-state";

import { schema } from "../../schema";
import { setMark, textFormattingToMarks, toggleUnderlineMark } from "./markUtils";

function applyCommand(state: EditorState, command: Command): EditorState {
  let nextState = state;
  command(state, (tr) => {
    nextState = nextState.apply(tr);
  });
  return nextState;
}

describe("rtl run direction round-trips through the mark helpers", () => {
  // eigenpal/docx-editor#806: the conversion path handled `rtl`, but the
  // live-edit / clipboard / keymap helpers (`textFormattingToMarks` /
  // `marksToTextFormatting`) had no rtl branch, so an Arabic/Hebrew run lost
  // its direction the moment it was re-marked in the editor.
  test("textFormattingToMarks emits the rtl mark", () => {
    const rtl = schema.marks.rtl;
    if (!rtl) {
      panic("Expected rtl mark in schema");
    }
    const marks = textFormattingToMarks({ rtl: true }, schema);
    expect(marks.some((mark) => mark.type === rtl)).toBe(true);
  });

  test("setMark(rtl) records rtl in the paragraph default text formatting", () => {
    const rtl = schema.marks.rtl;
    if (!rtl) {
      panic("Expected rtl mark in schema");
    }
    let state = EditorState.create({
      doc: schema.node("doc", null, [schema.node("paragraph")]),
      schema,
    });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1)));

    setMark(rtl, {})(state, (tr) => {
      state = state.apply(tr);
    });

    expect(state.doc.firstChild?.attrs.defaultTextFormatting).toEqual({
      rtl: true,
    });
  });
});

describe("mark commands", () => {
  test("preserve stored marks when setting formatting in an empty paragraph", () => {
    const fontSize = schema.marks.fontSize;
    if (!fontSize) {
      panic("Expected fontSize mark in schema");
    }

    let state = EditorState.create({
      doc: schema.node("doc", null, [schema.node("paragraph")]),
      schema,
    });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1)));

    const didSetMark = setMark(fontSize, { size: 32 })(state, (tr) => {
      state = state.apply(tr);
    });

    expect(didSetMark).toBe(true);
    expect(state.storedMarks?.some((mark) => mark.type === fontSize)).toBe(true);
    expect(state.doc.firstChild?.attrs.defaultTextFormatting).toEqual({
      fontSize: 32,
    });
  });

  test("selection font-family writes preserve untouched script slots and hint", () => {
    const fontFamily = schema.marks.fontFamily;
    if (!fontFamily) {
      panic("Expected fontFamily mark in schema");
    }

    let state = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("text", [
            fontFamily.create({
              ascii: "Arial",
              hAnsi: "Arial",
              eastAsia: "SimSun",
              cs: "Arial Unicode MS",
              hint: "eastAsia",
              asciiTheme: "majorHAnsi",
              hAnsiTheme: "majorHAnsi",
            }),
          ]),
        ]),
      ]),
      schema,
    });

    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 5)));
    state = applyCommand(state, setMark(fontFamily, { ascii: "Georgia", hAnsi: "Georgia" }));

    const mark = fontFamily.isInSet(state.doc.firstChild?.firstChild?.marks ?? []);
    expect(mark?.attrs).toMatchObject({
      ascii: "Georgia",
      hAnsi: "Georgia",
      eastAsia: "SimSun",
      cs: "Arial Unicode MS",
      hint: "eastAsia",
    });
    expect(mark?.attrs["asciiTheme"]).toBeNull();
    expect(mark?.attrs["hAnsiTheme"]).toBeNull();
  });

  test("font-family mark conversion preserves every script and theme slot", () => {
    const marks = textFormattingToMarks(
      {
        fontFamily: {
          ascii: "Arial",
          hAnsi: "Arial",
          eastAsia: "SimSun",
          cs: "Arial Unicode MS",
          hint: "eastAsia",
          asciiTheme: "majorAscii",
          hAnsiTheme: "majorHAnsi",
          eastAsiaTheme: "majorEastAsia",
          csTheme: "majorBidi",
        },
      },
      schema,
    );
    const mark = marks.find(({ type }) => type.name === "fontFamily");

    expect(mark?.attrs).toMatchObject({
      ascii: "Arial",
      hAnsi: "Arial",
      eastAsia: "SimSun",
      cs: "Arial Unicode MS",
      hint: "eastAsia",
      asciiTheme: "majorAscii",
      hAnsiTheme: "majorHAnsi",
      eastAsiaTheme: "majorEastAsia",
      csTheme: "majorBidi",
    });
  });

  test("collapsed-caret font-family writes preserve untouched script slots", () => {
    const fontFamily = schema.marks.fontFamily;
    if (!fontFamily) {
      panic("Expected fontFamily mark in schema");
    }

    let state = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("text", [
            fontFamily.create({
              ascii: "Arial",
              hAnsi: "Arial",
              eastAsia: "SimSun",
              cs: "Arial Unicode MS",
              hint: "eastAsia",
            }),
          ]),
        ]),
      ]),
      schema,
    });

    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3)));
    state = applyCommand(state, setMark(fontFamily, { ascii: "Verdana", hAnsi: "Verdana" }));

    const mark = fontFamily.isInSet(state.storedMarks ?? []);
    expect(mark?.attrs).toMatchObject({
      ascii: "Verdana",
      hAnsi: "Verdana",
      eastAsia: "SimSun",
      cs: "Arial Unicode MS",
      hint: "eastAsia",
    });
  });

  test("selection underline toggles preserve authored color across off and on", () => {
    const underline = schema.marks.underline;
    if (!underline) {
      panic("Expected underline mark in schema");
    }

    let state = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("text", [
            underline.create({
              style: "single",
              color: { rgb: "FF0000" },
            }),
          ]),
        ]),
      ]),
      schema,
    });

    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 5)));
    state = applyCommand(state, toggleUnderlineMark(underline));
    state = applyCommand(state, toggleUnderlineMark(underline));

    const mark = underline.isInSet(state.doc.firstChild?.firstChild?.marks ?? []);
    expect(mark?.attrs).toEqual({
      style: "single",
      color: { rgb: "FF0000" },
    });
  });

  test("collapsed-caret underline toggles preserve authored color across off and on", () => {
    const underline = schema.marks.underline;
    if (!underline) {
      panic("Expected underline mark in schema");
    }

    let state = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("text", [
            underline.create({
              style: "single",
              color: { rgb: "FF0000" },
            }),
          ]),
        ]),
      ]),
      schema,
    });

    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3)));
    state = applyCommand(state, toggleUnderlineMark(underline));
    state = applyCommand(state, toggleUnderlineMark(underline));

    const mark = underline.isInSet(state.storedMarks ?? []);
    expect(mark?.attrs).toEqual({
      style: "single",
      color: { rgb: "FF0000" },
    });
  });

  test("collapsed-caret underline toggles on when no underline mark is present", () => {
    const underline = schema.marks.underline;
    if (!underline) {
      panic("Expected underline mark in schema");
    }

    let state = EditorState.create({
      doc: schema.node("doc", null, [schema.node("paragraph", null, [schema.text("text")])]),
      schema,
    });

    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3)));
    state = applyCommand(state, toggleUnderlineMark(underline));

    const mark = underline.isInSet(state.storedMarks ?? []);
    expect(mark?.attrs).toMatchObject({ style: "single" });
  });

  test("underline none removes the visual decoration", () => {
    const underline = schema.marks.underline;
    if (!underline) {
      panic("Expected underline mark in schema");
    }

    const paragraph = schema.node("paragraph", null, [
      schema.text("text", [underline.create({ style: "none" })]),
    ]);
    const underlineMark = paragraph.child(0).marks.at(0);
    if (!underlineMark) {
      panic("Expected underline mark on text");
    }
    const domSpec = underlineMark.type.spec.toDOM?.(underlineMark);
    if (!Array.isArray(domSpec)) {
      panic("Expected underline mark to provide a DOM spec");
    }
    expect(domSpec[1]).toEqual({ style: "text-decoration: none" });
  });

  test("underline none survives the text-decoration parser", () => {
    const underline = schema.marks.underline;
    if (!underline) {
      panic("Expected underline mark in schema");
    }
    const parseRule = underline.spec.parseDOM?.find(
      (rule) => "style" in rule && rule.style === "text-decoration",
    );
    if (!parseRule || !("style" in parseRule) || !parseRule.getAttrs) {
      panic("Expected underline text-decoration parse rule");
    }

    expect(parseRule.getAttrs("none")).toEqual({ style: "none" });
    expect(parseRule.getAttrs("underline")).toEqual({});
  });
});
