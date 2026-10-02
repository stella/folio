import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";
import { getCanonicalCommandIntents, withCanonicalCommand } from "./canonicalCommands";
import { schema, singletonManager } from "./schema";

const stateWithSelection = () => {
  const doc = schema.node("doc", null, [
    schema.node("paragraph", { indentLeft: 360 }, [schema.text("first")]),
    schema.node("paragraph", { indentLeft: 720 }, [schema.text("second")]),
  ]);
  return EditorState.create({ schema, doc, selection: TextSelection.create(doc, 2, 10) });
};

describe("canonical command descriptors", () => {
  test("registers meaning without executing a command, including repeated probes", () => {
    const state = stateWithSelection();
    let executions = 0;
    const command = withCanonicalCommand(
      () => {
        executions += 1;
        return true;
      },
      () => [],
    );
    expect(getCanonicalCommandIntents(command, state)).toEqual([]);
    expect(getCanonicalCommandIntents(command, state)).toEqual([]);
    expect(executions).toBe(0);
    expect(getCanonicalCommandIntents(() => true, state)).toBeUndefined();
  });

  test("toolbar probes preserve document, selection and stored marks", () => {
    const state = stateWithSelection();
    const commands = [
      singletonManager.requireCommand("toggleBold")(),
      singletonManager.requireCommand("toggleItalic")(),
      singletonManager.requireCommand("toggleUnderline")(),
      singletonManager.requireCommand("setFontSize")(28),
      singletonManager.requireCommand("clearFontSize")(),
      singletonManager.requireCommand("setFontFamily")("Example Font"),
      singletonManager.requireCommand("clearFontFamily")(),
      singletonManager.requireCommand("setTextColor")({ rgb: "123456" }),
      singletonManager.requireCommand("clearTextColor")(),
      singletonManager.requireCommand("setHighlight")("yellow"),
      singletonManager.requireCommand("clearHighlight")(),
      singletonManager.requireCommand("alignCenter")(),
      singletonManager.requireCommand("setLineSpacing")(360),
      singletonManager.requireCommand("setSpaceBefore")(80),
      singletonManager.requireCommand("setSpaceAfter")(100),
      singletonManager.requireCommand("increaseIndent")(),
      singletonManager.requireCommand("decreaseIndent")(),
      singletonManager.requireCommand("setIndentLeft")(100),
      singletonManager.requireCommand("setIndentRight")(200),
      singletonManager.requireCommand("setIndentFirstLine")(300, true),
      singletonManager.requireCommand("applyStyle")("Heading1"),
      singletonManager.requireCommand("clearStyle")(),
      singletonManager.requireCommand("toggleBulletList")(),
      singletonManager.requireCommand("toggleNumberedList")(),
      singletonManager.requireCommand("increaseListLevel")(),
      singletonManager.requireCommand("decreaseListLevel")(),
      singletonManager.requireCommand("removeList")(),
      singletonManager.requireCommand("restartNumbering")(),
      singletonManager.requireCommand("continueNumbering")(),
      singletonManager.requireCommand("setNumberingValue")(7),
    ];
    const document = state.doc;
    const selection = state.selection;
    const storedMarks = state.storedMarks;
    for (const command of commands) {
      const first = getCanonicalCommandIntents(command, state);
      expect(first).toBeDefined();
      command(state);
      expect(getCanonicalCommandIntents(command, state)).toEqual(first);
      expect(state.doc).toBe(document);
      expect(state.selection).toBe(selection);
      expect(state.storedMarks).toBe(storedMarks);
    }
  });

  test("numbering command descriptors retain their start value", () => {
    const state = stateWithSelection();
    expect(
      getCanonicalCommandIntents(singletonManager.requireCommand("restartNumbering")(), state),
    ).toEqual([{ type: "restartNumbering" }]);
    expect(
      getCanonicalCommandIntents(singletonManager.requireCommand("setNumberingValue")(7), state),
    ).toEqual([{ type: "restartNumbering", start: 7 }]);
    expect(
      getCanonicalCommandIntents(singletonManager.requireCommand("continueNumbering")(), state),
    ).toEqual([{ type: "continueNumbering" }]);
  });

  test("formats each selected paragraph using its own indentation", () => {
    const state = stateWithSelection();
    expect(
      getCanonicalCommandIntents(singletonManager.requireCommand("increaseIndent")(120), state),
    ).toEqual([
      { type: "formatParagraph", at: 1, patch: { indentLeft: 480 } },
      { type: "formatParagraph", at: 8, patch: { indentLeft: 840 } },
    ]);
  });

  test("disabling styled toggle formatting states false for both script families", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("bold", [schema.mark("bold")])]),
    ]);
    const state = EditorState.create({ schema, doc, selection: TextSelection.create(doc, 1, 5) });
    expect(
      getCanonicalCommandIntents(singletonManager.requireCommand("toggleBold")(), state),
    ).toEqual([{ type: "formatRun", from: 1, to: 5, patch: { bold: false, boldCs: false } }]);
  });

  test("underline toggles preserve each selected run's underline color", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("a", [schema.mark("underline", { style: "single", color: { rgb: "123456" } })]),
        schema.text("b", [schema.mark("underline", { style: "double", color: { rgb: "654321" } })]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc, selection: TextSelection.create(doc, 1, 3) });
    expect(
      getCanonicalCommandIntents(singletonManager.requireCommand("toggleUnderline")(), state),
    ).toEqual([
      {
        type: "formatRun",
        from: 1,
        to: 2,
        patch: { underline: { style: "none", color: { rgb: "123456" } } },
      },
      {
        type: "formatRun",
        from: 2,
        to: 3,
        patch: { underline: { style: "none", color: { rgb: "654321" } } },
      },
    ]);
  });

  test("font-family writes retain each run's independent script-family properties", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("a", [schema.mark("fontFamily", { ascii: "Old", eastAsia: "East A" })]),
        schema.text("b", [schema.mark("fontFamily", { ascii: "Old", eastAsia: "East B" })]),
      ]),
    ]);
    const state = EditorState.create({ schema, doc, selection: TextSelection.create(doc, 1, 3) });
    expect(
      getCanonicalCommandIntents(singletonManager.requireCommand("setFontFamily")("New"), state),
    ).toEqual([
      {
        type: "formatRun",
        from: 1,
        to: 2,
        patch: { fontFamily: { ascii: "New", hAnsi: "New", eastAsia: "East A" } },
      },
      {
        type: "formatRun",
        from: 2,
        to: 3,
        patch: { fontFamily: { ascii: "New", hAnsi: "New", eastAsia: "East B" } },
      },
    ]);
  });
});
