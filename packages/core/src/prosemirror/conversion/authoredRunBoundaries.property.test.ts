import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { propertyConfig } from "../../../../../test/property-testing";
import { resolveWholeStory } from "../../internal/wholeStoryRevisionResolution";
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

  test("mixed text and note-reference content survives source-owned run joins", () => {
    fc.assert(
      fc.property(
        text,
        text,
        fc.constantFrom("footnoteRef", "endnoteRef"),
        fc.boolean(),
        (before, after, type, customMarkFollows) => {
          const content: Run[] = [
            {
              type: "run",
              content: [
                { type: "text", text: before },
                { type, id: 41, customMarkFollows },
                { type: "text", text: after },
              ],
            },
          ];
          const source = documentOf([{ type: "paragraph", content }]);
          const saved = fromProseDoc(toProseDoc(source), source);
          expect(savedRuns(saved)).toEqual(content);
          expect(savedRuns(fromProseDoc(toProseDoc(saved), saved))).toEqual(content);
        },
      ),
      propertyConfig(),
    );
  });

  test("an interior insertion extends one ordinary run and retains its authored neighbors", () => {
    const manager = new ExtensionManager(createStarterKit());
    manager.buildSchema();
    manager.initializeRuntime();
    fc.assert(
      fc.property(runs, fc.nat(), fc.nat(), (content, pickedRun, pickedOffset) => {
        const runIndex = pickedRun % content.length;
        const target = content.at(runIndex);
        if (target === undefined) throw new TypeError("Expected the generated run");
        const targetText = runText([target]);
        const offset = 1 + (pickedOffset % (targetText.length - 1));
        const position = 1 + runText(content.slice(0, runIndex)).length + offset;
        const source = documentOf([{ type: "paragraph", content }]);
        const state = EditorState.create({
          doc: toProseDoc(source),
          plugins: manager.getPlugins(),
        });
        const edited = state.apply(state.tr.insertText("!", position));
        const saved = savedRuns(fromProseDoc(edited.doc, source));
        expect(saved).toEqual(
          content.map((run, index) =>
            index === runIndex
              ? {
                  ...run,
                  content: [
                    {
                      type: "text",
                      text: targetText.slice(0, offset) + "!" + targetText.slice(offset),
                    },
                  ],
                }
              : run,
          ),
        );
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

  test("rejecting a reopened insertion rejoins its original run without joining neighbors", () => {
    fc.assert(
      fc.property(runs, fc.nat(), fc.boolean(), (content, pickedOffset, withSession) => {
        // SAFETY: the generator always produces at least two runs.
        const first = content[0]!;
        const original = withSession
          ? { ...first, preservedAttributes: [{ name: "rsidR", value: "00AB12CD" }] }
          : first;
        const originalText = runText([original]);
        const offset = 1 + (pickedOffset % (originalText.length - 1));
        const piece = (value: string): Run => ({
          ...original,
          content: [{ type: "text", text: value }],
        });
        const source = documentOf([
          {
            type: "paragraph",
            content: [
              piece(originalText.slice(0, offset)),
              {
                type: "insertion",
                info: { id: 4, author: "Reviewer" },
                content: [{ type: "run", content: [{ type: "text", text: "!" }] }],
              },
              piece(originalText.slice(offset)),
              ...content.slice(1),
            ],
          },
        ]);
        const { resolved } = resolveWholeStory({
          doc: toProseDoc(source),
          mode: "reject",
          styleResolver: null,
        });
        expect(savedRuns(fromProseDoc(resolved, source))).toEqual([original, ...content.slice(1)]);
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
