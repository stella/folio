import { afterAll, afterEach, beforeAll, describe, expect, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { Schema } from "prosemirror-model";
import fc from "fast-check";
import { EditorState, Plugin, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import { createSuggestionModePlugin } from "./plugins/suggestionMode";
import { createTextInputPlugin } from "./textInput";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: { content: "inline*", group: "block", toDOM: () => ["p", 0] },
    text: { group: "inline" },
  },
  marks: {
    strong: { toDOM: () => ["strong", 0] },
    insertion: {
      attrs: {
        revisionId: { default: 0 },
        author: { default: "" },
        date: { default: "" },
      },
      inclusive: false,
      toDOM: (mark) => ["ins", { "data-revision-id": String(mark.attrs["revisionId"]) }, 0],
    },
    deletion: {
      attrs: {
        revisionId: { default: 0 },
        author: { default: "" },
        date: { default: "" },
      },
      inclusive: false,
      toDOM: () => ["del", 0],
    },
  },
});

type CreateViewOptions = {
  editable?: boolean;
  plugins?: Plugin[];
  selection?: { from: number; to?: number };
};

const views: EditorView[] = [];

const createView = (
  content: Parameters<typeof schema.node>[2],
  options: CreateViewOptions = {},
) => {
  const doc = schema.node("doc", null, [schema.node("paragraph", null, content)]);
  const from = options.selection?.from ?? doc.content.size - 1;
  const to = options.selection?.to ?? from;
  const mount = document.createElement("div");
  document.body.append(mount);
  const view = new EditorView(mount, {
    state: EditorState.create({
      doc,
      plugins: [createTextInputPlugin(), ...(options.plugins ?? [])],
      selection: TextSelection.create(doc, from, to),
    }),
    dispatchTransaction(transaction) {
      view.updateState(view.state.apply(transaction));
    },
    editable: () => options.editable ?? true,
  });
  views.push(view);
  return view;
};

type BeforeInputOptions = {
  cancelable?: boolean;
  data?: string | null;
  inputType?: string;
  isComposing?: boolean;
  dataTransfer?: DataTransfer;
  targetRange?: AbstractRange;
};

const dispatchBeforeInput = (
  view: EditorView,
  {
    cancelable = true,
    data = "x",
    inputType = "insertText",
    isComposing = false,
    dataTransfer,
    targetRange,
  }: BeforeInputOptions = {},
) => {
  const event = new InputEvent("beforeinput", {
    bubbles: true,
    cancelable,
    data,
    dataTransfer,
    inputType,
    isComposing,
  });
  if (data === null) {
    Object.defineProperty(event, "data", { value: null });
  }
  if (dataTransfer) {
    Object.defineProperty(event, "dataTransfer", { value: dataTransfer });
  }
  if (inputType === "insertReplacementText") {
    Object.defineProperty(event, "getTargetRanges", {
      value: () => (targetRange ? [targetRange] : []),
    });
  }
  view.dom.dispatchEvent(event);
  return event;
};

const markedText = (view: EditorView, markName: string) => {
  let text = "";
  view.state.doc.descendants((node) => {
    if (node.isText && node.marks.some((mark) => mark.type.name === markName)) {
      text += node.text ?? "";
    }
  });
  return text;
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

describe("text input routing", () => {
  test.each([
    { markName: "insertion", edge: "before" },
    { markName: "insertion", edge: "after" },
    { markName: "deletion", edge: "before" },
    { markName: "deletion", edge: "after" },
  ])("preserves $markName contents at its $edge edge for Unicode input", ({ markName, edge }) => {
    const view = createView([schema.text("A")]);
    const text = fc
      .array(fc.constantFrom("a", "ž", "字", "😀", " ", "\u0301"), { minLength: 1, maxLength: 12 })
      .map((chars) => chars.join(""));
    fc.assert(
      fc.property(text, text, (revisionText, insertedText) => {
        const doc = schema.node("doc", null, [
          schema.node("paragraph", null, [
            schema.text("A"),
            schema.text(revisionText, [schema.mark(markName)]),
            schema.text("Z"),
          ]),
        ]);
        const position = edge === "before" ? 2 : revisionText.length + 2;
        view.updateState(
          EditorState.create({
            doc,
            plugins: view.state.plugins,
            selection: TextSelection.create(doc, position),
          }),
        );
        const event = dispatchBeforeInput(view, { data: insertedText });
        expect(event.defaultPrevented).toBe(true);
        expect(view.state.doc.textContent).toBe(
          edge === "before"
            ? `A${insertedText}${revisionText}Z`
            : `A${revisionText}${insertedText}Z`,
        );
        expect(markedText(view, markName)).toBe(revisionText);
      }),
      { numRuns: 40 },
    );
  });

  test("preserves stored marks for Unicode input", () => {
    const view = createView([schema.text("plain")]);
    view.dispatch(view.state.tr.setStoredMarks([schema.marks.strong.create()]));

    dispatchBeforeInput(view, { data: "ž🙂" });

    expect(view.state.doc.textContent).toBe("plainž🙂");
    expect(markedText(view, "strong")).toBe("ž🙂");
  });

  test("offers input to handleTextInput plugins before the default transaction", () => {
    const calls: Array<{ from: number; text: string; to: number }> = [];
    const handler = new Plugin({
      props: {
        handleTextInput(view, from, to, text) {
          calls.push({ from, text, to });
          view.dispatch(view.state.tr.insertText(`[${text}]`, from, to));
          return true;
        },
      },
    });
    const view = createView([schema.text("A")], { plugins: [handler] });

    dispatchBeforeInput(view, { data: "β" });

    expect(calls).toEqual([{ from: 2, text: "β", to: 2 }]);
    expect(view.state.doc.textContent).toBe("A[β]");
  });

  test("routes ordinary input through suggestion tracking", () => {
    const view = createView([schema.text("A")], {
      plugins: [createSuggestionModePlugin(true, "Reviewer")],
    });

    dispatchBeforeInput(view, { data: "新" });

    expect(view.state.doc.textContent).toBe("A新");
    expect(markedText(view, "insertion")).toBe("新");
  });

  test("preserves beforeinput tracking for standalone suggestion views", () => {
    const view = createView([schema.text("A")]);
    view.updateState(
      view.state.reconfigure({ plugins: [createSuggestionModePlugin(true, "Reviewer")] }),
    );
    const event = dispatchBeforeInput(view, { data: "新" });
    expect(event.defaultPrevented).toBe(true);
    expect(view.state.doc.textContent).toBe("A新");
    expect(markedText(view, "insertion")).toBe("新");
  });

  test("uses a replacement target range when it differs from the model selection", () => {
    const view = createView([schema.text("alpha beta omega")]);
    const textNode = view.dom.querySelector("p")?.firstChild;
    if (!textNode) throw new Error("expected rendered paragraph text");
    const targetRange = document.createRange();
    targetRange.setStart(textNode, 6);
    targetRange.setEnd(textNode, 10);

    const event = dispatchBeforeInput(view, {
      data: "BETA",
      inputType: "insertReplacementText",
      targetRange,
    });

    expect(event.defaultPrevented).toBe(true);
    expect(view.state.doc.textContent).toBe("alpha BETA omega");
  });

  test("does not intercept through the composition-end microtask", async () => {
    const view = createView([schema.text("A")]);
    view.dom.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true }));

    const composingEvent = dispatchBeforeInput(view, { data: "あ" });
    view.dom.dispatchEvent(new CompositionEvent("compositionend", { bubbles: true }));
    const settlingEvent = dispatchBeforeInput(view, { data: "い" });

    expect(composingEvent.defaultPrevented).toBe(false);
    expect(settlingEvent.defaultPrevented).toBe(false);
    expect(view.state.doc.textContent).toBe("A");

    await Promise.resolve();
    const settledEvent = dispatchBeforeInput(view, { data: "う" });
    expect(settledEvent.defaultPrevented).toBe(true);
    expect(view.state.doc.textContent).toBe("Aう");
  });

  test("leaves replacement input without exactly one target range to native handling", () => {
    const view = createView([schema.text("alpha")]);

    const event = dispatchBeforeInput(view, {
      data: "ALPHA",
      inputType: "insertReplacementText",
    });

    expect(event.defaultPrevented).toBe(false);
    expect(view.state.doc.textContent).toBe("alpha");
  });

  test("reads text/plain replacement data when InputEvent.data is null", () => {
    const view = createView([schema.text("alpha beta")]);
    const textNode = view.dom.querySelector("p")?.firstChild;
    if (!textNode) throw new Error("expected rendered paragraph text");
    const dataTransfer = new DataTransfer();
    dataTransfer.setData("text/plain", "BETA");
    const targetRange = document.createRange();
    targetRange.setStart(textNode, 6);
    targetRange.setEnd(textNode, 10);

    const event = dispatchBeforeInput(view, {
      data: null,
      dataTransfer,
      inputType: "insertReplacementText",
      targetRange,
    });

    expect(event.defaultPrevented).toBe(true);
    expect(view.state.doc.textContent).toBe("alpha BETA");
  });

  test("allows an empty replacement to delete its target range", () => {
    const view = createView([schema.text("AB")]);
    const textNode = view.dom.querySelector("p")?.firstChild;
    if (!textNode) throw new Error("expected rendered paragraph text");
    const targetRange = document.createRange();
    targetRange.setStart(textNode, 1);
    targetRange.setEnd(textNode, 2);

    const event = dispatchBeforeInput(view, {
      data: "",
      inputType: "insertReplacementText",
      targetRange,
    });

    expect(event.defaultPrevented).toBe(true);
    expect(view.state.doc.textContent).toBe("A");
  });

  test.each([
    { name: "read-only", options: { editable: false }, event: {} },
    { name: "non-cancelable", options: {}, event: { cancelable: false } },
    { name: "composition event", options: {}, event: { isComposing: true } },
    { name: "unsupported input type", options: {}, event: { inputType: "insertParagraph" } },
  ])("leaves $name input to the browser", ({ options, event: eventOptions }) => {
    const view = createView([schema.text("A")], options);

    const event = dispatchBeforeInput(view, eventOptions);

    expect(event.defaultPrevented).toBe(false);
    expect(view.state.doc.textContent).toBe("A");
  });
});
