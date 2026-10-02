/**
 * Each `LISTNUM` field a list marker draws exists once, whatever is done
 * around it. A split leaves it in one half, a join carries it into the joined
 * paragraph, typing moves it along with its neighbours, and a paste brings
 * none with it. So after any sequence of those, the document holds the fields
 * it was opened with, in the order it was opened with, and each of them once.
 *
 * Every field in the fixture caches a display of its own, so one written
 * twice, or one that changed places with another, shows.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { removeFoldedListNumberFields } from "../prosemirror/extensions/features/pasteCleanup";
import {
  bodyParagraphs,
  foldedCaptureNodes,
  inlineTokens,
  listNumberFieldDocx,
  openDocx,
  type ParagraphSpec,
} from "./__tests__/listNumberFieldFixture";
import { isFoldedListNumberCapture } from "./foldedListNumberFields";
import { serializeParagraph } from "./serializer/paragraphSerializer";

setDefaultTimeout(propertyTestTimeout(30_000));

const SPECS: ParagraphSpec[] = [
  {
    paraId: "20000001",
    marker: "decimal",
    fields: [
      {
        instruction: " LISTNUM ",
        result: "(a)",
        formatting: "plain",
        before: "",
        gap: [],
        tab: true,
      },
      {
        instruction: "LISTNUM",
        result: "(b)",
        formatting: "bold",
        before: "one ",
        gap: [],
        tab: false,
      },
    ],
    body: "first body",
  },
  {
    paraId: "20000002",
    marker: "decimal",
    fields: [
      {
        instruction: " LISTNUM ",
        result: "(c)",
        formatting: "plain",
        before: "",
        gap: [],
        tab: true,
      },
    ],
    body: "second body",
  },
  {
    paraId: "20000003",
    marker: "percent",
    fields: [
      {
        instruction: " LISTNUM ",
        result: "(d)",
        formatting: "plain",
        before: "two ",
        gap: [],
        tab: true,
      },
    ],
    body: "third body",
  },
];

const RESULTS = ["text:(a)", "text:(b)", "text:(c)", "text:(d)"];

/** A place in the document: a paragraph, and how far into its content. */
const placeArbitrary = fc.record({
  paragraph: fc.nat({ max: 999 }).map((thousandths) => thousandths / 1000),
  offset: fc.nat({ max: 1000 }).map((thousandths) => thousandths / 1000),
});

type Place = { paragraph: number; offset: number };

const stepArbitrary = fc.oneof(
  fc.record({ kind: fc.constant("split" as const), at: placeArbitrary }),
  fc.record({ kind: fc.constant("join" as const), at: placeArbitrary }),
  fc.record({ kind: fc.constant("type" as const), at: placeArbitrary }),
  fc.record({
    kind: fc.constant("paste" as const),
    from: placeArbitrary,
    to: placeArbitrary,
    at: placeArbitrary,
  }),
);

type Step =
  | { kind: "split" | "join" | "type"; at: Place }
  | {
      kind: "paste";
      from: Place;
      to: Place;
      at: Place;
    };

const paragraphsOf = (doc: PMNode): { node: PMNode; position: number }[] => {
  const paragraphs: { node: PMNode; position: number }[] = [];
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  doc.forEach((node, position) => {
    if (node.type.name === "paragraph") {
      paragraphs.push({ node, position });
    }
  });
  return paragraphs;
};

/** The document position `place` names, inside a paragraph's content. */
const positionOf = (doc: PMNode, place: Place): number => {
  const paragraphs = paragraphsOf(doc);
  const paragraph = paragraphs[Math.floor(place.paragraph * paragraphs.length)];
  if (!paragraph) {
    throw new Error("The document holds no paragraph");
  }
  return paragraph.position + 1 + Math.round(place.offset * paragraph.node.content.size);
};

const applyStep = (state: EditorState, step: Step): EditorState => {
  switch (step.kind) {
    case "split":
      return state.apply(state.tr.split(positionOf(state.doc, step.at)));
    case "join": {
      const paragraphs = paragraphsOf(state.doc);
      const first = paragraphs[Math.floor(step.at.paragraph * (paragraphs.length - 1))];
      // One paragraph has nothing to join.
      if (!first || paragraphs.length < 2) {
        return state;
      }
      return state.apply(state.tr.join(first.position + first.node.nodeSize));
    }
    case "type":
      return state.apply(state.tr.insertText("x", positionOf(state.doc, step.at)));
    case "paste": {
      const a = positionOf(state.doc, step.from);
      const b = positionOf(state.doc, step.to);
      const copied = state.doc.slice(Math.min(a, b), Math.max(a, b));
      const at = positionOf(state.doc, step.at);
      // What the editor's paste does to a slice before it lands.
      return state.apply(state.tr.replaceRange(at, at, removeFoldedListNumberFields(copied)));
    }
    default: {
      const unhandled: never = step;
      throw new Error(`Unhandled step ${JSON.stringify(unhandled)}`);
    }
  }
};

const captureMarkup = (doc: PMNode): unknown[] =>
  foldedCaptureNodes(doc).map(({ node }) => [node.attrs["foldedListNumber"], node.attrs["xml"]]);

describe("the LISTNUM fields of a document under splits, joins, typing and pastes", () => {
  test(
    "each field is still held once, and in the order the document was opened with",
    async () => {
      const parsed = await openDocx(await listNumberFieldDocx(SPECS));
      const opened = EditorState.create({ doc: toProseDoc(parsed) });
      const original = captureMarkup(opened.doc);
      expect(original).toHaveLength(7);

      assertProperty(
        fc.property(fc.array(stepArbitrary, { minLength: 1, maxLength: 12 }), (steps) => {
          let state = opened;
          for (const step of steps) {
            state = applyStep(state, step);
            expect(captureMarkup(state.doc)).toEqual(original);
          }

          const paragraphs = bodyParagraphs(fromProseDoc(state.doc, parsed));
          const captures = paragraphs.flatMap((paragraph) =>
            paragraph.content.filter(isFoldedListNumberCapture),
          );
          expect(captures.map(({ foldedListNumber, xml }) => [foldedListNumber, xml])).toEqual(
            original,
          );
          // The same count read off the markup a save writes.
          const tokens = paragraphs.flatMap((paragraph) =>
            inlineTokens(serializeParagraph(paragraph)),
          );
          expect(tokens.filter((token) => RESULTS.includes(token))).toEqual(RESULTS);
          expect(tokens.filter((token) => token === "fldChar:begin")).toHaveLength(4);
        }),
        { numRuns: 150 },
      );
    },
    propertyTestTimeout(120_000),
  );
});
