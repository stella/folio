import { afterAll, afterEach, beforeAll, describe, expect, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { Fragment, Schema, Slice } from "prosemirror-model";
import { ReplaceStep } from "prosemirror-transform";
import { EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import { createCanonicalInputBoundary } from "./canonicalInput";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: { content: "inline*", group: "block", toDOM: () => ["p", 0] },
    text: { group: "inline" },
  },
  marks: { strong: { toDOM: () => ["strong", 0] } },
});

const views: EditorView[] = [];

const createRig = (from = 2, to = from) => {
  const inputs: { from: number; to: number; text: string }[] = [];
  const refusals: string[] = [];
  const history: string[] = [];
  const boundary = createCanonicalInputBoundary({
    replace: (input) => inputs.push(input),
    refuse: (reason) => refusals.push(reason),
    undo: () => {
      history.push("undo");
      return true;
    },
    redo: () => {
      history.push("redo");
      return true;
    },
  });
  const doc = schema.node("doc", null, [schema.node("paragraph", null, schema.text("A😀B"))]);
  const mount = document.createElement("div");
  document.body.append(mount);
  const view = new EditorView(mount, {
    state: EditorState.create({ doc, selection: TextSelection.create(doc, from, to) }),
    handleTextInput: boundary.handleTextInput,
    handleKeyDown: boundary.handleKeyDown,
    handleDOMEvents: boundary.handleDOMEvents,
    dispatchTransaction(transaction) {
      if (transaction.docChanged) {
        if (!boundary.commitNativeProposal(view, transaction)) boundary.refuseNativeMutation(view);
        view.updateState(view.state);
        return;
      }
      view.updateState(view.state.apply(transaction));
    },
  });
  views.push(view);
  return { boundary, view, inputs, refusals, history };
};

beforeAll(() => GlobalRegistrator.register());
afterEach(() => {
  for (const view of views.splice(0)) {
    const mount = view.dom.parentElement;
    view.destroy();
    mount?.remove();
  }
  document.body.replaceChildren();
});
afterAll(() => GlobalRegistrator.unregister());

describe("canonical input boundary", () => {
  test("cancelable beforeinput emits one classified intent without a PM edit", () => {
    const { boundary, view, inputs } = createRig(2, 4);
    const original = view.state;
    const event = new InputEvent("beforeinput", {
      inputType: "insertText",
      data: "𐐀",
      cancelable: true,
    });
    expect(boundary.handleDOMEvents.beforeinput(view, event)).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(inputs).toEqual([{ from: 2, to: 4, text: "𐐀" }]);
    expect(view.state).toBe(original);
  });

  test("cancelable beforeinput authorization expires before any later native flush", () => {
    const { boundary, view, inputs, refusals } = createRig();
    boundary.handleDOMEvents.beforeinput(
      view,
      new InputEvent("beforeinput", { inputType: "insertText", data: "a", cancelable: true }),
    );
    boundary.handleTextInput(view, 2, 2, "b");
    expect(inputs).toEqual([{ from: 2, to: 2, text: "a" }]);
    expect(refusals).toHaveLength(1);
  });

  test("an unaddressable cancelable replacement is refused before its DOM mutation", () => {
    const { boundary, view, inputs, refusals } = createRig();
    const event = new InputEvent("beforeinput", {
      inputType: "insertReplacementText",
      data: "x",
      cancelable: true,
    });
    expect(boundary.handleDOMEvents.beforeinput(view, event)).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    boundary.handleTextInput(view, 2, 2, "x");
    expect(inputs).toEqual([]);
    expect(refusals).toHaveLength(1);
  });

  test.each([
    { key: "Backspace", from: 4, to: 4 },
    { key: "Delete", from: 2, to: 2 },
    { key: "Delete", from: 2, to: 4 },
  ])("$key deletes complete code points or the selection", ({ key, from, to }) => {
    const { boundary, view, inputs } = createRig(from, to);
    const event = new KeyboardEvent("keydown", { key, cancelable: true });
    expect(boundary.handleKeyDown(view, event)).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(inputs).toEqual([{ from: 2, to: 4, text: "" }]);
    expect(view.state.doc.textContent).toBe("A😀B");
  });

  test.each(["deleteContentBackward", "deleteContentForward"])(
    "%s beforeinput emits one complete-character deletion",
    (inputType) => {
      const { boundary, view, inputs } = createRig(inputType.endsWith("Backward") ? 4 : 2);
      const event = new InputEvent("beforeinput", { inputType, cancelable: true });
      expect(boundary.handleDOMEvents.beforeinput(view, event)).toBe(true);
      expect(event.defaultPrevented).toBe(true);
      expect(inputs).toEqual([{ from: 2, to: 4, text: "" }]);
    },
  );

  test("a native proposal is accepted exactly once and only at its originating state", () => {
    const { boundary, view, inputs } = createRig(2, 4);
    const propose = () =>
      boundary.handleDOMEvents.beforeinput(
        view,
        new InputEvent("beforeinput", {
          inputType: "insertText",
          data: "x",
          cancelable: false,
        }),
      );
    expect(propose()).toBe(false);
    const transaction = view.state.tr.insertText("x", 2, 4);
    expect(boundary.takeNativeProposal(view.state, transaction)).toEqual({
      from: 2,
      to: 4,
      text: "x",
    });
    expect(boundary.takeNativeProposal(view.state, transaction)).toBeNull();
    expect(inputs).toEqual([]);
    propose();
    const previous = view.state;
    view.updateState(
      view.state.apply(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1))),
    );
    expect(boundary.takeNativeProposal(view.state, previous.tr.insertText("x", 2, 4))).toBeNull();
  });

  test.each(["text", "range", "mark", "extraStep", "structure", "structureFlag"])(
    "a classified native insertion rejects a forged %s proposal",
    (change) => {
      const { boundary, view } = createRig(2, change === "structureFlag" ? 2 : 4);
      boundary.handleDOMEvents.beforeinput(
        view,
        new InputEvent("beforeinput", {
          inputType: "insertText",
          data: "x",
          cancelable: false,
        }),
      );
      const transaction = view.state.tr;
      switch (change) {
        case "text":
          transaction.insertText("y", 2, 4);
          break;
        case "range":
          transaction.insertText("x", 1, 4);
          break;
        case "mark":
          transaction.addStoredMark(schema.mark("strong")).insertText("x", 2, 4);
          break;
        case "extraStep":
          transaction.insertText("x", 2, 4).insertText("z", 1);
          break;
        case "structure":
          transaction.replaceWith(2, 4, schema.node("paragraph"));
          break;
        case "structureFlag":
          transaction.step(
            new ReplaceStep(2, 2, new Slice(Fragment.from(schema.text("x")), 0, 0), true),
          );
          break;
      }
      expect(boundary.takeNativeProposal(view.state, transaction)).toBeNull();
      expect(view.state.doc.textContent).toBe("A😀B");
    },
  );

  test("input flush expires an unused native proposal", async () => {
    const { boundary, view } = createRig();
    boundary.handleDOMEvents.beforeinput(
      view,
      new InputEvent("beforeinput", {
        inputType: "insertText",
        data: "x",
        cancelable: false,
      }),
    );
    boundary.handleDOMEvents.input(view);
    await Promise.resolve();
    expect(boundary.takeNativeProposal(view.state, view.state.tr.insertText("x", 2))).toBeNull();
  });

  test("composition rejects provisional and final inputs through the end microtask", async () => {
    const { boundary, view, inputs, refusals } = createRig();
    boundary.handleDOMEvents.compositionstart(
      view,
      new Event("compositionstart", { cancelable: true }),
    );
    boundary.handleTextInput(view, 2, 2, "契");
    const event = new InputEvent("beforeinput", {
      inputType: "insertText",
      data: "契",
      cancelable: true,
      isComposing: true,
    });
    expect(boundary.handleDOMEvents.beforeinput(view, event)).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    boundary.handleDOMEvents.compositionend(view);
    boundary.handleTextInput(view, 2, 2, "契約");
    expect(inputs).toEqual([]);
    expect(refusals).toHaveLength(1);
    await Promise.resolve();
    boundary.handleDOMEvents.beforeinput(
      view,
      new InputEvent("beforeinput", { inputType: "insertText", data: "a", cancelable: true }),
    );
    expect(inputs).toEqual([{ from: 2, to: 2, text: "a" }]);
  });

  test.each(["foreign", "replacement", "composition"])(
    "%s DOM text cannot commit through handleTextInput or a native transaction",
    async (kind) => {
      const { boundary, view, inputs, refusals } = createRig();
      const original = view.state;
      if (kind === "composition") {
        boundary.handleDOMEvents.compositionstart(view, new Event("compositionstart"));
        boundary.handleDOMEvents.compositionend(view);
        await Promise.resolve();
      }
      if (kind !== "foreign") {
        const event = new InputEvent("beforeinput", {
          inputType: kind === "composition" ? "insertFromComposition" : "insertReplacementText",
          data: "x",
          cancelable: false,
        });
        expect(boundary.handleDOMEvents.beforeinput(view, event)).toBe(true);
        expect(event.defaultPrevented).toBe(false);
      }
      const textNode = view.dom.querySelector("p")?.firstChild;
      expect(textNode).toBeDefined();
      if (textNode) textNode.textContent = "Ax😀B";
      expect(boundary.handleTextInput(view, 2, 2, "x")).toBe(true);
      expect(boundary.commitNativeProposal(view, original.tr.insertText("x", 2))).toBe(false);
      boundary.handleDOMEvents.input(view);
      await Promise.resolve();
      // The observer can flush again after the input microtask; it still lacks authorization.
      boundary.handleTextInput(view, 2, 2, "x");
      expect(inputs).toEqual([]);
      expect(view.state).toBe(original);
      expect(refusals.length).toBeLessThanOrEqual(2);
    },
  );

  test.each(["state", "from", "to", "text", "match"])(
    "native handleTextInput requires an exact originating %s proposal",
    (change) => {
      const { boundary, view, inputs } = createRig(2, 4);
      boundary.handleDOMEvents.beforeinput(
        view,
        new InputEvent("beforeinput", { inputType: "insertText", data: "x" }),
      );
      if (change === "state") {
        view.updateState(
          view.state.apply(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1))),
        );
      }
      boundary.handleTextInput(
        view,
        change === "from" ? 1 : 2,
        change === "to" ? 2 : 4,
        change === "text" ? "y" : "x",
      );
      expect(inputs).toEqual(change === "match" ? [{ from: 2, to: 4, text: "x" }] : []);
      boundary.handleTextInput(view, 2, 4, "x");
      expect(inputs).toHaveLength(change === "match" ? 1 : 0);
    },
  );

  test("native transaction commits a classified proposal once", () => {
    const { boundary, view, inputs } = createRig(2, 4);
    boundary.handleDOMEvents.beforeinput(
      view,
      new InputEvent("beforeinput", { inputType: "insertText", data: "x" }),
    );
    const transaction = view.state.tr.insertText("x", 2, 4);
    expect(boundary.commitNativeProposal(view, transaction)).toBe(true);
    expect(boundary.commitNativeProposal(view, transaction)).toBe(false);
    expect(inputs).toEqual([{ from: 2, to: 4, text: "x" }]);
  });

  test.each(["blur", "beforeinput"])(
    "a missing compositionend recovers on %s without repeated refusals",
    async (recovery) => {
      const { boundary, view, inputs, refusals } = createRig();
      boundary.handleDOMEvents.compositionstart(view, new Event("compositionstart"));
      for (let index = 0; index < 5; index += 1) {
        boundary.handleTextInput(view, 2, 2, "契");
        boundary.handleDOMEvents.beforeinput(
          view,
          new InputEvent("beforeinput", { inputType: "insertCompositionText", data: "契" }),
        );
        boundary.handleDOMEvents.input(view);
        await Promise.resolve();
      }
      expect(refusals).toHaveLength(1);
      expect(inputs).toEqual([]);
      if (recovery === "blur") boundary.handleDOMEvents.blur(view);
      boundary.handleDOMEvents.beforeinput(
        view,
        new InputEvent("beforeinput", { inputType: "insertText", data: "a", cancelable: true }),
      );
      expect(inputs).toEqual([{ from: 2, to: 2, text: "a" }]);
      boundary.handleDOMEvents.compositionstart(view, new Event("compositionstart"));
      expect(refusals).toHaveLength(2);
    },
  );

  test.each(["insertParagraph", "insertFromPaste", "deleteByCut", "insertFromDrop"])(
    "%s is refused without an intent",
    (inputType) => {
      const { boundary, view, inputs, refusals } = createRig();
      const event = new InputEvent("beforeinput", { inputType, cancelable: true });
      expect(boundary.handleDOMEvents.beforeinput(view, event)).toBe(true);
      expect(event.defaultPrevented).toBe(true);
      expect(inputs).toEqual([]);
      expect(refusals).toHaveLength(1);
    },
  );

  test.each(["paste", "cut", "drop"] as const)(
    "%s DOM proposals are refused atomically",
    (kind) => {
      const { boundary, view, inputs, refusals } = createRig();
      const event = new Event(kind, { cancelable: true });
      expect(boundary.handleDOMEvents[kind](view, event)).toBe(true);
      expect(event.defaultPrevented).toBe(true);
      expect(inputs).toEqual([]);
      expect(refusals).toHaveLength(1);
      expect(view.state.doc.textContent).toBe("A😀B");
    },
  );

  test.each([
    { key: "Backspace", from: 1, ctrlKey: false },
    { key: "Delete", from: 5, ctrlKey: false },
    { key: "Backspace", from: 3, ctrlKey: false },
    { key: "Backspace", from: 4, ctrlKey: true },
  ])("unsafe deletion at $from is refused without an intent", ({ key, from, ctrlKey }) => {
    const { boundary, view, inputs, refusals } = createRig(from);
    const event = new KeyboardEvent("keydown", { key, ctrlKey, cancelable: true });
    expect(boundary.handleKeyDown(view, event)).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(inputs).toEqual([]);
    expect(refusals).toHaveLength(1);
    expect(view.state.doc.textContent).toBe("A😀B");
  });

  test("keyboard and beforeinput history routes share the semantic journal", () => {
    const { boundary, view, history, inputs } = createRig();
    for (const options of [
      { key: "z", ctrlKey: true },
      { key: "z", metaKey: true, shiftKey: true },
      { key: "y", ctrlKey: true },
    ]) {
      const event = new KeyboardEvent("keydown", { ...options, cancelable: true });
      expect(boundary.handleKeyDown(view, event)).toBe(true);
      expect(event.defaultPrevented).toBe(true);
    }
    for (const inputType of ["historyUndo", "historyRedo"]) {
      const event = new InputEvent("beforeinput", { inputType, cancelable: true });
      expect(boundary.handleDOMEvents.beforeinput(view, event)).toBe(true);
      expect(event.defaultPrevented).toBe(true);
    }
    expect(history).toEqual(["undo", "redo", "redo", "undo", "redo"]);
    expect(inputs).toEqual([]);
  });
});
