import {
  afterAll,
  afterEach,
  beforeAll,
  describe,
  expect,
  setDefaultTimeout,
  test,
} from "bun:test";
import fc from "fast-check";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { Fragment, Schema, Slice } from "prosemirror-model";
import { ReplaceStep } from "prosemirror-transform";
import { EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";

import { createCanonicalInputBoundary } from "./canonicalInput";

setDefaultTimeout(propertyTestTimeout(30_000));

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: { content: "inline*", group: "block", toDOM: () => ["p", 0] },
    text: { group: "inline" },
  },
  marks: { strong: { toDOM: () => ["strong", 0] } },
});

type CanonicalReplacement = Parameters<
  Parameters<typeof createCanonicalInputBoundary>[0]["replace"]
>[0];

const views: EditorView[] = [];

const createRig = (from = 2, to = from) => {
  const inputs: CanonicalReplacement[] = [];
  const compositionTransitions: string[] = [];
  const refusals: string[] = [];
  const history: string[] = [];
  const groupBoundaries: number[] = [];
  const boundary = createCanonicalInputBoundary({
    replace: (input) => inputs.push(input),
    breakUndoGroup: () => {
      groupBoundaries.push(1);
    },
    beginComposition: () => {
      compositionTransitions.push("begin");
      return true;
    },
    endComposition: () => compositionTransitions.push("end"),
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
      if (boundary.acceptComposition(view, transaction)) return;
      if (transaction.docChanged) {
        if (!boundary.commitNativeProposal(view, transaction)) boundary.refuseNativeMutation(view);
        view.updateState(view.state);
        return;
      }
      view.updateState(view.state.apply(transaction));
    },
  });
  views.push(view);
  return { boundary, view, inputs, refusals, history, compositionTransitions, groupBoundaries };
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
  test.each(["keydown", "mousedown", "blur", "paste", "cut", "drop", "compositionstart"] as const)(
    "%s expires the preceding gesture's native proposal",
    (gesture) => {
      const { boundary, view, inputs } = createRig();
      boundary.handleDOMEvents.beforeinput(
        view,
        new InputEvent("beforeinput", {
          inputType: "insertText",
          data: "x",
        }),
      );
      const proposed = view.state.tr.insertText("x", 2);
      if (gesture === "keydown")
        boundary.handleKeyDown(view, new KeyboardEvent("keydown", { key: "ArrowLeft" }));
      else if (gesture === "paste" || gesture === "cut" || gesture === "drop")
        boundary.handleDOMEvents[gesture](view, new Event(gesture, { cancelable: true }));
      else boundary.handleDOMEvents[gesture](view);
      expect(boundary.takeNativeProposal(view.state, proposed)).toBeNull();
      expect(inputs).toEqual([]);
      boundary.reset();
    },
  );

  test("each refused keyboard gesture surfaces its own refusal without a browser input event", () => {
    const { boundary, view, refusals, groupBoundaries } = createRig();
    for (let index = 0; index < 3; index++) {
      boundary.handleKeyDown(
        view,
        new KeyboardEvent("keydown", { key: "Enter", cancelable: true }),
      );
    }
    expect(refusals).toHaveLength(3);
    expect(groupBoundaries).toHaveLength(3);
  });

  test("typing gestures preserve runs; navigation and refused clipboard gestures close them", () => {
    const { boundary, view, groupBoundaries } = createRig();
    for (const key of ["a", "b", "Backspace", "Delete"]) {
      boundary.handleKeyDown(view, new KeyboardEvent("keydown", { key }));
    }
    expect(groupBoundaries).toEqual([]);
    for (const key of ["ArrowLeft", "Home", "Tab"]) {
      boundary.handleKeyDown(view, new KeyboardEvent("keydown", { key }));
    }
    expect(groupBoundaries).toHaveLength(3);
    boundary.handleDOMEvents.paste(view, new Event("paste", { cancelable: true }));
    boundary.handleDOMEvents.cut(view, new Event("cut", { cancelable: true }));
    boundary.handleDOMEvents.drop(view, new Event("drop", { cancelable: true }));
    expect(groupBoundaries).toHaveLength(6);
  });

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
    expect(inputs).toEqual([{ from: 2, to: 4, text: "𐐀", semantic: "typing" }]);
    expect(view.state).toBe(original);
  });

  test("cancelable beforeinput authorization expires before any later native flush", () => {
    const { boundary, view, inputs, refusals } = createRig();
    boundary.handleDOMEvents.beforeinput(
      view,
      new InputEvent("beforeinput", { inputType: "insertText", data: "a", cancelable: true }),
    );
    boundary.handleTextInput(view, 2, 2, "b");
    expect(inputs).toEqual([{ from: 2, to: 2, text: "a", semantic: "typing" }]);
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
    expect(inputs).toEqual([
      {
        from: 2,
        to: 4,
        text: "",
        semantic: key === "Backspace" ? "deleteBackward" : "deleteForward",
      },
    ]);
    expect(view.state.doc.textContent).toBe("A😀B");
  });

  test.each(["deleteContentBackward", "deleteContentForward"])(
    "%s beforeinput emits one complete-character deletion",
    (inputType) => {
      const { boundary, view, inputs } = createRig(inputType.endsWith("Backward") ? 4 : 2);
      const event = new InputEvent("beforeinput", { inputType, cancelable: true });
      expect(boundary.handleDOMEvents.beforeinput(view, event)).toBe(true);
      expect(event.defaultPrevented).toBe(true);
      expect(inputs).toEqual([
        {
          from: 2,
          to: 4,
          text: "",
          semantic: inputType.endsWith("Backward") ? "deleteBackward" : "deleteForward",
        },
      ]);
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
      semantic: "typing",
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

  test.each([
    { from: 2, to: 2, final: "契約", expectedTo: 2 },
    { from: 2, to: 4, final: "契約", expectedTo: 4 },
  ])(
    "CJK composition at $from..$to lowers one net replacement",
    async ({ from, to, final, expectedTo }) => {
      const { boundary, view, inputs, refusals, compositionTransitions } = createRig(from, to);
      const baseline = view.state;
      expect(boundary.handleDOMEvents.compositionstart(view)).toBe(false);
      expect(boundary.isComposing).toBe(true);
      expect(boundary.handleTextInput(view, from, to, "契")).toBe(false);
      const provisional = new InputEvent("beforeinput", {
        inputType: "insertCompositionText",
        data: "契",
        cancelable: true,
        isComposing: true,
      });
      expect(boundary.handleDOMEvents.beforeinput(view, provisional)).toBe(false);
      expect(provisional.defaultPrevented).toBe(false);
      view.dispatch(view.state.tr.insertText("契", from, to).setMeta("composition", 1));
      view.dispatch(view.state.tr.insertText(final, from, from + 1).setMeta("composition", 1));
      expect(view.state.doc.textContent).toBe(from === to ? "A契約😀B" : "A契約B");
      expect(inputs).toEqual([]);
      expect(boundary.handleDOMEvents.compositionend(view)).toBe(false);
      await new Promise<void>((resolve) => setTimeout(resolve, 40));
      expect(inputs).toEqual([{ from, to: expectedTo, text: final, semantic: "composition" }]);
      expect(view.state).toBe(baseline);
      expect(boundary.isComposing).toBe(false);
      expect(compositionTransitions).toEqual(["begin", "end"]);
      expect(refusals).toEqual([]);
    },
  );

  test("late final flush after the end microtask and duplicate events commit once", async () => {
    const { boundary, view, inputs, refusals, compositionTransitions } = createRig(2, 4);
    boundary.handleDOMEvents.compositionstart(view);
    boundary.handleDOMEvents.compositionstart(view);
    view.dispatch(view.state.tr.insertText("契", 2, 4).setMeta("composition", 1));
    boundary.handleDOMEvents.compositionend(view);
    boundary.handleDOMEvents.compositionend(view);
    await Promise.resolve();
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    expect(inputs).toEqual([]);
    const final = new InputEvent("beforeinput", {
      inputType: "insertFromComposition",
      data: "契約",
      cancelable: true,
    });
    expect(boundary.handleDOMEvents.beforeinput(view, final)).toBe(false);
    expect(final.defaultPrevented).toBe(false);
    expect(boundary.handleTextInput(view, 2, 3, "契約")).toBe(false);
    view.dispatch(view.state.tr.insertText("契約", 2, 3).setMeta("composition", 1));
    boundary.handleDOMEvents.input(view);
    boundary.handleDOMEvents.input(view);
    await new Promise<void>((resolve) => setTimeout(resolve, 40));
    expect(inputs).toEqual([{ from: 2, to: 4, text: "契約", semantic: "composition" }]);
    expect(compositionTransitions).toEqual(["begin", "end"]);
    expect(refusals).toEqual([]);
  });

  test("classified late final input authorizes one flush after PM clears its composition metadata", async () => {
    const { boundary, view, inputs, refusals } = createRig(2, 4);
    boundary.handleDOMEvents.compositionstart(view);
    view.dispatch(view.state.tr.insertText("契", 2, 4).setMeta("composition", 1));
    boundary.handleDOMEvents.compositionend(view);
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    const final = new InputEvent("beforeinput", {
      inputType: "insertFromComposition",
      data: "契約",
      cancelable: true,
    });
    expect(boundary.handleDOMEvents.beforeinput(view, final)).toBe(false);
    view.dispatch(view.state.tr.insertText("契約", 2, 3));
    expect(view.state.doc.textContent).toBe("A契約B");
    view.dispatch(view.state.tr.insertText("foreign", 1));
    expect(view.state.doc.textContent).toBe("A契約B");
    expect(inputs).toEqual([]);
    expect(refusals).toHaveLength(1);
    boundary.handleDOMEvents.input(view);
    await new Promise<void>((resolve) => setTimeout(resolve, 40));
    expect(inputs).toEqual([{ from: 2, to: 4, text: "契約", semantic: "composition" }]);
  });

  test("Escape cancels provisional composition and restores its exact baseline", async () => {
    const { boundary, view, inputs, refusals, compositionTransitions } = createRig(2, 4);
    const baseline = view.state;
    boundary.handleDOMEvents.compositionstart(view);
    view.dispatch(view.state.tr.insertText("契約", 2, 4).setMeta("composition", 1));
    boundary.handleDOMEvents.compositionend(view);
    const escape = new KeyboardEvent("keydown", { key: "Escape", cancelable: true });
    expect(boundary.handleKeyDown(view, escape)).toBe(true);
    expect(escape.defaultPrevented).toBe(true);
    await new Promise<void>((resolve) => setTimeout(resolve, 40));
    expect(view.state).toBe(baseline);
    expect(inputs).toEqual([]);
    expect(refusals).toEqual([]);
    expect(compositionTransitions).toEqual(["begin", "end"]);
  });

  test("other transactions and history shortcuts wait for composition", () => {
    const { boundary, view, inputs, refusals, history } = createRig(2, 4);
    boundary.handleDOMEvents.compositionstart(view);
    view.dispatch(view.state.tr.insertText("契", 2, 4).setMeta("composition", 1));
    const provisional = view.state;
    view.dispatch(view.state.tr.insertText("foreign", 1));
    expect(view.state).toBe(provisional);
    expect(boundary.isComposing).toBe(true);
    expect(refusals).toHaveLength(1);
    const undo = new KeyboardEvent("keydown", { key: "z", ctrlKey: true, cancelable: true });
    expect(boundary.handleKeyDown(view, undo)).toBe(true);
    expect(undo.defaultPrevented).toBe(true);
    expect(history).toEqual([]);
    expect(inputs).toEqual([]);
    boundary.handleKeyDown(view, new KeyboardEvent("keydown", { key: "Escape" }));
  });

  test("composition without any net change leaves the journal empty", async () => {
    const { boundary, view, inputs } = createRig(2, 4);
    const baseline = view.state;
    boundary.handleDOMEvents.compositionstart(view);
    view.dispatch(view.state.tr.insertText("契", 2, 4).setMeta("composition", 1));
    view.dispatch(view.state.tr.insertText("😀", 2, 3).setMeta("composition", 1));
    boundary.handleDOMEvents.compositionend(view);
    await new Promise<void>((resolve) => setTimeout(resolve, 40));
    expect(inputs).toEqual([]);
    expect(view.state).toBe(baseline);
  });

  test.each(["mark", "structure", "foreignParagraph"])(
    "composition refuses a %s mutation atomically",
    (change) => {
      const { boundary, view, inputs, refusals } = createRig(2, 4);
      const baseline = view.state;
      boundary.handleDOMEvents.compositionstart(view);
      const transaction = view.state.tr;
      if (change === "mark")
        transaction.addStoredMark(schema.mark("strong")).insertText("契", 2, 4);
      else if (change === "structure") transaction.replaceWith(2, 4, schema.node("paragraph"));
      else
        transaction.insert(
          transaction.doc.content.size,
          schema.node("paragraph", null, schema.text("foreign")),
        );
      view.dispatch(transaction.setMeta("composition", 1));
      expect(view.state).toBe(baseline);
      expect(inputs).toEqual([]);
      expect(refusals).toHaveLength(1);
      expect(boundary.isComposing).toBe(true);
      // A later native flush of the refused gesture must not become a new edit.
      view.dispatch(view.state.tr.insertText("late", 2, 4).setMeta("composition", 1));
      expect(view.state).toBe(baseline);
      expect(inputs).toEqual([]);
      expect(refusals).toHaveLength(1);
      boundary.handleKeyDown(view, new KeyboardEvent("keydown", { key: "Escape" }));
      expect(boundary.isComposing).toBe(false);
    },
  );

  test("generated cross-paragraph compositions lower one exact replacement or cancel", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 0, max: 4 }),
        fc.integer({ min: 1, max: 4 }),
        fc.array(fc.constantFrom("契", "約", "😀", "alpha"), { minLength: 1, maxLength: 4 }),
        fc.boolean(),
        async (start, end, updates, cancel) => {
          const { boundary, view, inputs, refusals, compositionTransitions } = createRig();
          const doc = schema.node("doc", null, [
            schema.node("paragraph", null, schema.text("Alpha")),
            schema.node("paragraph", null, schema.text("Beta")),
          ]);
          const from = 1 + start;
          const to = 8 + end;
          view.updateState(
            EditorState.create({ doc, selection: TextSelection.create(doc, from, to) }),
          );
          const baseline = view.state;
          boundary.handleDOMEvents.compositionstart(view);
          let width = to - from;
          for (const text of updates) {
            view.dispatch(
              view.state.tr.insertText(text, from, from + width).setMeta("composition", 1),
            );
            width = text.length;
          }
          if (cancel) boundary.handleKeyDown(view, new KeyboardEvent("keydown", { key: "Escape" }));
          else boundary.handleDOMEvents.compositionend(view);
          await new Promise<void>((resolve) => setTimeout(resolve, 40));
          expect(inputs).toEqual(
            cancel ? [] : [{ from, to, text: updates.at(-1), semantic: "composition" }],
          );
          expect(view.state).toBe(baseline);
          expect(refusals).toEqual([]);
          expect(compositionTransitions).toEqual(["begin", "end"]);
          view.destroy();
          views.splice(views.indexOf(view), 1);
          view.dom.parentElement?.remove();
        },
      ),
      { numRuns: 50 },
    );
  });

  test("generated composition event traces lower one final replacement or cancel", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.record({
          replacement: fc.boolean(),
          provisional: fc.array(fc.constantFrom("契", "約", "文", "書", "😀"), {
            minLength: 1,
            maxLength: 5,
          }),
          final: fc.array(fc.constantFrom("契", "約", "文", "書"), { minLength: 1, maxLength: 5 }),
          duplicates: fc.integer({ min: 0, max: 3 }),
          finalOrder: fc.constantFrom("beforeEnd", "microtask", "nextTask"),
          completion: fc.constantFrom("commit", "cancel", "blur", "nextInput"),
        }),
        async ({ replacement, provisional, final, duplicates, finalOrder, completion }) => {
          const { boundary, view, inputs, refusals, compositionTransitions } = createRig(
            2,
            replacement ? 4 : 2,
          );
          const baseline = view.state;
          boundary.handleDOMEvents.compositionstart(view);
          let width = replacement ? 2 : 0;
          for (const text of provisional) {
            view.dispatch(view.state.tr.insertText(text, 2, 2 + width).setMeta("composition", 1));
            width = text.length;
          }
          const text = final.join("");
          const flushFinal = () => {
            const event = new InputEvent("beforeinput", {
              inputType: "insertFromComposition",
              data: text,
              cancelable: true,
            });
            expect(boundary.handleDOMEvents.beforeinput(view, event)).toBe(false);
            expect(event.defaultPrevented).toBe(false);
            view.dispatch(view.state.tr.insertText(text, 2, 2 + width).setMeta("composition", 1));
            boundary.handleDOMEvents.input(view);
          };
          if (finalOrder === "beforeEnd") flushFinal();
          boundary.handleDOMEvents.compositionend(view);
          if (finalOrder === "microtask") {
            await Promise.resolve();
            flushFinal();
          }
          if (finalOrder === "nextTask") {
            await new Promise<void>((resolve) => setTimeout(resolve, 0));
            flushFinal();
          }
          for (let index = 0; index < duplicates; index++) {
            boundary.handleDOMEvents.compositionend(view);
            boundary.handleDOMEvents.input(view);
          }
          expect(inputs).toEqual([]);
          if (completion === "cancel") {
            boundary.handleKeyDown(view, new KeyboardEvent("keydown", { key: "Escape" }));
          } else if (completion === "blur") boundary.handleDOMEvents.blur(view);
          else if (completion === "nextInput")
            boundary.handleDOMEvents.beforeinput(
              view,
              new InputEvent("beforeinput", {
                inputType: "insertText",
                data: "a",
                cancelable: true,
              }),
            );
          await new Promise<void>((resolve) => setTimeout(resolve, 40));
          const expected: CanonicalReplacement[] =
            completion === "cancel"
              ? []
              : [{ from: 2, to: replacement ? 4 : 2, text, semantic: "composition" }];
          if (completion === "nextInput")
            expected.push({ from: 2, to: replacement ? 4 : 2, text: "a", semantic: "typing" });
          expect(inputs).toEqual(expected);
          expect(view.state).toBe(baseline);
          expect(boundary.isComposing).toBe(false);
          expect(compositionTransitions).toEqual(["begin", "end"]);
          expect(refusals).toEqual([]);
          view.destroy();
          views.splice(views.indexOf(view), 1);
          view.dom.parentElement?.remove();
        },
      ),
      { numRuns: 50 },
    );
  });

  test.each(["foreign", "replacement", "composition"])(
    "%s DOM text cannot commit through handleTextInput or a native transaction",
    async (kind) => {
      const { boundary, view, inputs, refusals } = createRig();
      const original = view.state;
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
      expect(inputs).toEqual(
        change === "match" ? [{ from: 2, to: 4, text: "x", semantic: "typing" }] : [],
      );
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
    expect(inputs).toEqual([{ from: 2, to: 4, text: "x", semantic: "typing" }]);
  });

  test.each(["blur", "beforeinput"])(
    "a missing compositionend commits its net replacement on %s",
    (recovery) => {
      const { boundary, view, inputs, refusals, compositionTransitions } = createRig(2, 4);
      const baseline = view.state;
      view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      expect(view.composing).toBe(true);
      view.dispatch(view.state.tr.insertText("契", 2, 4).setMeta("composition", 1));
      view.dispatch(view.state.tr.insertText("契約", 2, 3).setMeta("composition", 1));
      expect(inputs).toEqual([]);
      if (recovery === "blur") boundary.handleDOMEvents.blur(view);
      else
        boundary.handleDOMEvents.beforeinput(
          view,
          new InputEvent("beforeinput", { inputType: "insertText", data: "a", cancelable: true }),
        );
      expect(inputs).toEqual([
        { from: 2, to: 4, text: "契約", semantic: "composition" },
        ...(recovery === "beforeinput" ? [{ from: 2, to: 4, text: "a", semantic: "typing" }] : []),
      ]);
      expect(view.state).toBe(baseline);
      expect(boundary.isComposing).toBe(false);
      expect(compositionTransitions).toEqual(["begin", "end"]);
      expect(refusals).toEqual([]);
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
