import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { registerClipboardIntentHandler } from "../clipboardIntent";

import {
  buildPlainTextSlice,
  CLIPBOARD_READ_ERROR_EVENT,
  pasteWithoutFormatting,
} from "./pastePlainText";

// Minimal block schema (no DOM needed to construct or fill it).
const blockSchema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: { group: "block", content: "inline*" },
    text: { group: "inline" },
  },
  marks: {
    bold: { toDOM: () => ["strong", 0], parseDOM: [{ tag: "strong" }] },
  },
});

function paragraphTexts(fragment: PMNode["content"]): string[] {
  const texts: string[] = [];
  fragment.forEach((node) => texts.push(node.textContent));
  return texts;
}

function everyTextNodeIsUnmarked(fragment: PMNode["content"]): boolean {
  let clean = true;
  fragment.descendants((node) => {
    if (node.isText && node.marks.length > 0) {
      clean = false;
    }
    return true;
  });
  return clean;
}

describe("buildPlainTextSlice", () => {
  test("single line becomes one paragraph of unmarked text", () => {
    const slice = buildPlainTextSlice("hello world", blockSchema);
    expect(paragraphTexts(slice.content)).toEqual(["hello world"]);
    expect(everyTextNodeIsUnmarked(slice.content)).toBe(true);
  });

  test("newline runs split into one paragraph each and collapse blank lines", () => {
    const slice = buildPlainTextSlice("a\nb\n\nc", blockSchema);
    expect(paragraphTexts(slice.content)).toEqual(["a", "b", "c"]);
  });

  test("carriage returns are treated as paragraph breaks", () => {
    const slice = buildPlainTextSlice("x\r\ny", blockSchema);
    expect(paragraphTexts(slice.content)).toEqual(["x", "y"]);
  });

  test("opens both ends so text merges into the surrounding block", () => {
    const slice = buildPlainTextSlice("merge me", blockSchema);
    expect(slice.openStart).toBe(1);
    expect(slice.openEnd).toBe(1);
  });

  test("carries no source formatting even when text looks like markup", () => {
    const slice = buildPlainTextSlice("**not bold** <b>either</b>", blockSchema);
    expect(paragraphTexts(slice.content)).toEqual(["**not bold** <b>either</b>"]);
    expect(everyTextNodeIsUnmarked(slice.content)).toBe(true);
  });

  test("falls back to flat text when the schema has no paragraph node", () => {
    const inlineSchema = new Schema({
      nodes: {
        doc: { content: "text*" },
        text: {},
      },
    });
    const slice = buildPlainTextSlice("just text", inlineSchema);
    expect(slice.content.textBetween(0, slice.content.size)).toBe("just text");
  });
});

describe("pasteWithoutFormatting dry run", () => {
  const withClipboard = (readText: (() => Promise<string>) | undefined, run: () => void): void => {
    const original = Object.getOwnPropertyDescriptor(globalThis, "navigator");
    Object.defineProperty(globalThis, "navigator", {
      configurable: true,
      value: { clipboard: readText ? { readText } : undefined },
    });
    try {
      run();
    } finally {
      if (original) {
        Object.defineProperty(globalThis, "navigator", original);
      } else {
        Reflect.deleteProperty(globalThis, "navigator");
      }
    }
  };

  const fakeState = { schema: blockSchema } as unknown as EditorState;

  test("a dispatch-less probe reports available without touching the clipboard", () => {
    let read = false;
    withClipboard(
      () => {
        read = true;
        return Promise.resolve("x");
      },
      () => {
        expect(pasteWithoutFormatting(fakeState)).toBe(true);
        expect(read).toBe(false);
      },
    );
  });

  test("reports unavailable when the runtime has no clipboard reader", () => {
    withClipboard(undefined, () => {
      expect(pasteWithoutFormatting(fakeState)).toBe(false);
    });
  });

  test("emits a clipboard-read-error event when the read is denied", async () => {
    const dom = new EventTarget();
    let firedError: unknown;
    dom.addEventListener(CLIPBOARD_READ_ERROR_EVENT, (event) => {
      firedError = (event as CustomEvent).detail?.error;
    });
    const view = {
      dom,
      isDestroyed: false,
      state: fakeState,
      dispatch: () => undefined,
    } as unknown as EditorView;

    withClipboard(
      () => Promise.reject(new Error("denied")),
      () => {
        expect(pasteWithoutFormatting(fakeState, () => undefined, view)).toBe(true);
      },
    );

    // Flush the rejected-read microtasks so the catch handler runs.
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(firedError).toBeInstanceOf(Error);
  });

  test("an asynchronous clipboard read uses the current state and owning intent handler", async () => {
    const original = EditorState.create({
      schema: blockSchema,
      doc: blockSchema.node(
        "doc",
        null,
        blockSchema.node("paragraph", null, blockSchema.text("abcd")),
      ),
    });
    const read = Promise.withResolvers<string>();
    let nativeDispatches = 0;
    const view = {
      dom: new EventTarget(),
      isDestroyed: false,
      state: original,
      dispatch: () => {
        nativeDispatches += 1;
      },
    } as unknown as EditorView;
    const handled: {
      plain: boolean;
      from: number;
      to: number;
      texts: string[];
      unmarked: boolean;
    }[] = [];
    registerClipboardIntentHandler(view, (slice, plain) => {
      handled.push({
        plain,
        from: view.state.selection.from,
        to: view.state.selection.to,
        texts: paragraphTexts(slice.content),
        unmarked: everyTextNodeIsUnmarked(slice.content),
      });
      view.state = view.state.apply(view.state.tr.replaceSelection(slice));
      return true;
    });
    withClipboard(
      () => read.promise,
      () => {
        expect(pasteWithoutFormatting(original, () => undefined, view)).toBe(true);
      },
    );
    // The read belongs to the current view, including intervening edits and selection movement.
    view.state = original.apply(original.tr.insertText("Z", 1));
    view.state = view.state.apply(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 4, 5)),
    );
    read.resolve("X\r\nY");
    await read.promise;
    await Promise.resolve();
    expect(handled).toEqual([{ plain: true, from: 4, to: 5, texts: ["X", "Y"], unmarked: true }]);
    expect(paragraphTexts(view.state.doc.content)).toEqual(["ZabX", "Yd"]);
    expect(nativeDispatches).toBe(0);
  });

  test.each(["refused", "destroyed"] as const)(
    "a %s asynchronous clipboard intent cannot fall through to native mutation",
    async (status) => {
      const read = Promise.withResolvers<string>();
      let nativeDispatches = 0;
      let intentCalls = 0;
      const view = {
        dom: new EventTarget(),
        isDestroyed: false,
        state: fakeState,
        dispatch: () => {
          nativeDispatches += 1;
        },
      } as unknown as EditorView;
      registerClipboardIntentHandler(view, () => {
        intentCalls += 1;
        return false;
      });
      withClipboard(
        () => read.promise,
        () => {
          expect(pasteWithoutFormatting(fakeState, () => undefined, view)).toBe(true);
        },
      );
      if (status === "destroyed") Object.defineProperty(view, "isDestroyed", { value: true });
      read.resolve("pasted");
      await read.promise;
      await Promise.resolve();
      expect(intentCalls).toBe(status === "destroyed" ? 0 : 1);
      expect(nativeDispatches).toBe(0);
    },
  );
});
