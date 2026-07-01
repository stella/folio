import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";

import { buildPlainTextSlice, pasteWithoutFormatting } from "./pastePlainText";

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
});
