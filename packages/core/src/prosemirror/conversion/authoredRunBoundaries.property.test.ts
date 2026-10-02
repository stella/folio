import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { propertyConfig } from "../../../../../test/property-testing";
import type { Document, Paragraph, Run } from "../../types/document";
import { createStarterKit } from "../extensions/StarterKit";
import { ExtensionManager } from "../extensions/ExtensionManager";
import { fromProseDoc, proseDocToBlocks } from "./fromProseDoc";
import { headerFooterToProseDoc, toProseDoc } from "./toProseDoc";

const text = fc.stringMatching(/^[a-z]{2,8}$/u);
const formatting = fc.constantFrom(
  undefined,
  { bold: false },
  { noProof: true },
  { fontSizeCs: 24 },
);
const runs = fc
  .tuple(fc.array(text, { minLength: 2, maxLength: 6 }), formatting)
  .map(([texts, authored]) =>
    texts.map(
      (value): Run => ({
        type: "run",
        content: [{ type: "text", text: value }],
        ...(authored === undefined ? {} : { formatting: authored }),
      }),
    ),
  );

const runText = (content: Run[]): string =>
  content
    .flatMap((run) => run.content)
    .filter((item) => item.type === "text")
    .map((item) => item.text)
    .join("");

const documentOf = (content: Paragraph[]): Document => ({ package: { document: { content } } });

const savedRuns = (document: Document): Run[] => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new TypeError("Expected the synthetic paragraph");
  return paragraph.content.filter((item): item is Run => item.type === "run");
};

// The previous generator exercised only runs with attribute/sink payloads.
// Ordinary runs also own boundaries, even when their marks are identical.
describe("authored run boundary ownership", () => {
  test("no-op and repeated saves preserve every nonempty authored run", () => {
    fc.assert(
      fc.property(runs, (content) => {
        const source = documentOf([{ type: "paragraph", content }]);
        const saved = fromProseDoc(toProseDoc(source), source);
        expect(savedRuns(saved)).toEqual(content);
        expect(savedRuns(fromProseDoc(toProseDoc(saved), saved))).toEqual(content);
      }),
      propertyConfig(),
    );
  });

  test("an insertion retains untouched neighboring runs through the production plugins", () => {
    const manager = new ExtensionManager(createStarterKit());
    manager.buildSchema();
    manager.initializeRuntime();
    fc.assert(
      fc.property(runs, (content) => {
        const source = documentOf([{ type: "paragraph", content }]);
        const state = EditorState.create({
          doc: toProseDoc(source),
          plugins: manager.getPlugins(),
        });
        const edited = state.apply(state.tr.insertText("!", 2));
        const saved = savedRuns(fromProseDoc(edited.doc, source));
        const untouched = content.slice(1);
        expect(saved.slice(-untouched.length)).toEqual(untouched);
        expect(runText(saved)).toBe(runText(content).replace(/^./u, "$&!"));
      }),
      propertyConfig(),
    );
  });

  test("joining paragraphs cannot collide their authored run identities", () => {
    fc.assert(
      fc.property(runs, (content) => {
        const source = documentOf(content.map((run) => ({ type: "paragraph", content: [run] })));
        let state = EditorState.create({ doc: toProseDoc(source) });
        while (state.doc.childCount > 1) {
          state = state.apply(state.tr.join(state.doc.child(0).nodeSize));
        }
        expect(savedRuns(fromProseDoc(state.doc, source))).toEqual(content);
      }),
      propertyConfig(),
    );
  });

  test("secondary stories preserve the same authored boundaries", () => {
    fc.assert(
      fc.property(runs, (content) => {
        expect(proseDocToBlocks(headerFooterToProseDoc([{ type: "paragraph", content }]))).toEqual([
          { type: "paragraph", content },
        ]);
      }),
      propertyConfig(),
    );
  });
});
