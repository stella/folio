import { expect, test } from "bun:test";
import type { Node } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

import { schema } from "../prosemirror/schema";
import { handleReadOnlySelectionKeyDown } from "./readOnlySelection";

const selectionHarness = (text: string) => {
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [schema.text(text)]),
  ]);
  return stateHarness(doc, 1);
};

const stateHarness = (doc: Node, position: number) => {
  let state = EditorState.create({
    doc,
    selection: TextSelection.create(doc, position),
  });
  const view = {
    get state() {
      return state;
    },
    dispatch: (transaction: Parameters<typeof state.apply>[0]) => {
      state = state.apply(transaction);
    },
  };
  return {
    view,
    get state() {
      return state;
    },
  };
};

const keyEvent = (
  key: string,
  overrides: Partial<
    Pick<KeyboardEvent, "altKey" | "ctrlKey" | "metaKey" | "shiftKey">
  > = {},
) => {
  let prevented = false;
  return {
    get defaultPrevented() {
      return prevented;
    },
    altKey: overrides.altKey ?? false,
    ctrlKey: overrides.ctrlKey ?? false,
    key,
    metaKey: overrides.metaKey ?? false,
    preventDefault: () => {
      prevented = true;
    },
    shiftKey: overrides.shiftKey ?? false,
  };
};

test("read-only horizontal Shift+Arrow advances by grapheme boundaries", () => {
  const text = "A👩‍👩‍👧‍👦e\u0301B";
  const harness = selectionHarness(text);
  const selected = () =>
    harness.state.doc.textBetween(
      harness.state.selection.from,
      harness.state.selection.to,
    );

  for (const expected of ["A", "A👩‍👩‍👧‍👦", "A👩‍👩‍👧‍👦e\u0301"]) {
    const event = keyEvent("ArrowRight", { shiftKey: true });
    expect(handleReadOnlySelectionKeyDown(harness.view, event)).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(selected()).toBe(expected);
  }
});

test("read-only selection leaves native and modified caret keys unclaimed", () => {
  const harness = selectionHarness("text");
  for (const event of [
    keyEvent("ArrowRight"),
    keyEvent("ArrowRight", { shiftKey: true, ctrlKey: true }),
    keyEvent("ArrowRight", { shiftKey: true, metaKey: true }),
    keyEvent("ArrowRight", { shiftKey: true, altKey: true }),
  ]) {
    expect(handleReadOnlySelectionKeyDown(harness.view, event)).toBe(false);
    expect(event.defaultPrevented).toBe(false);
    expect(harness.state.selection.empty).toBe(true);
  }
});

test("read-only Shift+Arrow crosses an inline atom and paragraph boundary", () => {
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.text("A"),
      schema.node("hardBreak"),
      schema.text("B"),
    ]),
    schema.node("paragraph", null, [schema.text("C")]),
  ]);
  const inlineHarness = stateHarness(doc, 2);
  const inlineEvent = keyEvent("ArrowRight", { shiftKey: true });
  expect(handleReadOnlySelectionKeyDown(inlineHarness.view, inlineEvent)).toBe(
    true,
  );
  expect(inlineHarness.state.selection.from).toBe(2);
  expect(inlineHarness.state.selection.to).toBe(3);
  expect(inlineHarness.state.doc.nodeAt(2)?.type.name).toBe("hardBreak");

  const paragraphEnd = stateHarness(doc, 4);
  const paragraphEvent = keyEvent("ArrowRight", { shiftKey: true });
  expect(
    handleReadOnlySelectionKeyDown(paragraphEnd.view, paragraphEvent),
  ).toBe(true);
  expect(paragraphEnd.state.selection.from).toBe(4);
  expect(paragraphEnd.state.selection.to).toBe(6);
});

test("read-only horizontal selection leaves RTL visual arrows native", () => {
  const doc = schema.node("doc", null, [
    schema.node(
      "paragraph",
      { direction: { source: "manual", value: "rtl" } },
      [schema.text("אב")],
    ),
  ]);
  const harness = stateHarness(doc, 1);
  const event = keyEvent("ArrowRight", { shiftKey: true });
  expect(handleReadOnlySelectionKeyDown(harness.view, event)).toBe(false);
  expect(event.defaultPrevented).toBe(false);
  expect(harness.state.selection.empty).toBe(true);
});
