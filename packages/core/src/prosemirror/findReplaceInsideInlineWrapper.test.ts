/**
 * Find and replace reach the text a transparent inline wrapper holds.
 *
 * The model-side projection was an `if`-chain that answered "no text" for
 * every member it did not name, and `inlineWrapper` was one of them: a phrase
 * inside a `w:bdo`, a `w:dir`, a smart tag or a run-level `w:customXml` could
 * be read on the page and not found. The revision wrappers were in the same
 * class, which is why the fix is a `switch` with a `never` default rather than
 * a case per kind.
 *
 * Both halves are asserted together because they have to agree: the model
 * projection produces the paragraph offsets and the editor resolves them back
 * to a position, so a projection that skips content shifts every later offset
 * and the replacement lands on the wrong characters.
 */

import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { createDefaultFindOptions, findInDocument } from "../utils/findReplace";
import { createEmptyDocument } from "../utils/createDocument";
import type { Document, InlineWrapper, Paragraph, ParagraphContent } from "../types/document";
import { fromProseDoc } from "./conversion/fromProseDoc";
import { toProseDoc } from "./conversion/toProseDoc";
import { resolveFindMatchRange } from "./findReplaceSelection";
import { schema } from "./schema";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

/**
 * The wrapper holds text on both sides of the phrase.
 *
 * The mark is `inclusive: false`, so replacing every character a wrapper
 * covers legitimately removes it: the author replaced the whole of what the
 * tag named. What must survive is a replacement *inside* the span, which is
 * what a find-and-replace across a document does.
 */
const WRAPPED_TEXT = "the stock price";
const BEFORE = "before ";
const AFTER = " after";
const MATCH_START = BEFORE.length + "the ".length;

const WRAPPERS: ReadonlyArray<{ name: string; wrapper: InlineWrapper }> = [
  {
    name: "a bidirectional override",
    wrapper: {
      type: "inlineWrapper",
      kind: "bidi",
      control: "override",
      direction: "rtl",
      content: [run(WRAPPED_TEXT)],
    },
  },
  {
    name: "a smart tag",
    wrapper: {
      type: "inlineWrapper",
      kind: "smartTag",
      element: "City",
      uri: "urn:example:tags",
      content: [run(WRAPPED_TEXT)],
    },
  },
  {
    name: "a custom-XML wrapper",
    wrapper: {
      type: "inlineWrapper",
      kind: "customXml",
      element: "party",
      content: [run(WRAPPED_TEXT)],
    },
  },
  {
    name: "an insertion",
    wrapper: {
      type: "inlineWrapper",
      kind: "bidi",
      control: "embedding",
      content: [
        {
          type: "insertion",
          info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
          content: [run(WRAPPED_TEXT)],
        },
      ],
    },
  },
  {
    name: "a smart tag inside a bidirectional embedding",
    wrapper: {
      type: "inlineWrapper",
      kind: "bidi",
      control: "embedding",
      direction: "rtl",
      content: [
        {
          type: "inlineWrapper",
          kind: "smartTag",
          element: "City",
          content: [run(WRAPPED_TEXT)],
        },
      ],
    },
  },
];

const documentWith = (content: Paragraph["content"]): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [{ type: "paragraph", paraId: "C0000001", content }],
      },
    },
  };
};

/** The outermost wrappers only, so nested ones do not count their text twice. */
const wrappersIn = (content: readonly ParagraphContent[]): InlineWrapper[] => {
  const found: InlineWrapper[] = [];
  const visit = (items: readonly ParagraphContent[]): void => {
    for (const item of items) {
      if (item.type === "inlineWrapper") {
        found.push(item);
        continue;
      }
      if (
        item.type === "insertion" ||
        item.type === "deletion" ||
        item.type === "moveFrom" ||
        item.type === "moveTo"
      ) {
        visit(item.content);
      }
    }
  };
  visit(content);
  return found;
};

const textIn = (content: readonly ParagraphContent[]): string => {
  let text = "";
  const visit = (items: readonly ParagraphContent[]): void => {
    for (const item of items) {
      switch (item.type) {
        case "run":
          for (const part of item.content) {
            if (part.type === "text") {
              text += part.text;
            }
          }
          break;
        case "inlineWrapper":
        case "insertion":
        case "deletion":
        case "moveFrom":
        case "moveTo":
          visit(item.content);
          break;
        default:
          break;
      }
    }
  };
  visit(content);
  return text;
};

describe("the offsets the model projection produces", () => {
  test("agree with the editor's own searchable text", () => {
    // Deleted text is struck through on the page and is a text node in the
    // editor, so the model projection has to count it too: a member counted on
    // one side and not the other shifts every later offset, and the
    // replacement lands on the wrong characters.
    const source = documentWith([
      run("keep "),
      {
        type: "deletion",
        info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
        content: [run("gone ")],
      },
      run("the stock price"),
    ]);
    const match = findInDocument(source, "stock", createDefaultFindOptions()).at(0);
    if (match === undefined) {
      throw new Error("The phrase was not found");
    }
    const state = EditorState.create({ schema, doc: toProseDoc(source) });
    const range = resolveFindMatchRange(state.doc, match);
    if (range === null) {
      throw new Error("The editor could not resolve the match");
    }
    expect(state.doc.textBetween(range.from, range.to)).toBe("stock");
  });
});

describe("a phrase inside a transparent inline wrapper", () => {
  for (const { name, wrapper } of WRAPPERS) {
    test(`is found inside ${name}`, () => {
      const source = documentWith([run(BEFORE), wrapper, run(AFTER)]);
      expect(
        findInDocument(source, "stock", createDefaultFindOptions()).map(
          ({ startOffset, endOffset, text }) => ({ startOffset, endOffset, text }),
        ),
      ).toEqual([{ startOffset: MATCH_START, endOffset: MATCH_START + 5, text: "stock" }]);
    });

    test(`is replaced inside ${name}, and the wrapper survives`, () => {
      const source = documentWith([run(BEFORE), wrapper, run(AFTER)]);
      const match = findInDocument(source, "stock", createDefaultFindOptions()).at(0);
      if (match === undefined) {
        throw new Error("The phrase was not found");
      }
      const state = EditorState.create({ schema, doc: toProseDoc(source) });
      const range = resolveFindMatchRange(state.doc, match);
      if (range === null) {
        throw new Error("The editor could not resolve the match");
      }
      const replaced = state.apply(state.tr.insertText("shares", range.from, range.to));
      const block = fromProseDoc(replaced.doc, source).package.document.content.at(0);
      if (block?.type !== "paragraph") {
        throw new Error("The replacement lost its paragraph");
      }
      expect(textIn(block.content)).toBe(`${BEFORE}the shares price${AFTER}`);
      expect(textIn(wrappersIn(block.content))).toBe("the shares price");
    });
  }
});
