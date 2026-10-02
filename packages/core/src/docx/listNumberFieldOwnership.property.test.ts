/**
 * Under any sequence of splits, joins, typing and pastes, the editor never
 * hides a `LISTNUM` field the page does not show, and never hides one twice.
 *
 * After every step each paragraph is in the form the fold allows: its
 * captures open it and its marker shows exactly their fields. A capture is
 * one of those the document was opened with, and stands no more often than it did. And at the end
 * the fields each paragraph shows are the fields its saved markup holds, with
 * every field the document was opened with still among them.
 *
 * Every field in the fixture caches a display of its own, so one written
 * twice, hidden, or lost shows.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, type Transaction } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import {
  foldedListNumberPlugin,
  unfoldPastedListNumberFields,
} from "../prosemirror/foldedListNumber";
import { schema } from "../prosemirror/schema";
import {
  bodyParagraphs,
  editorState,
  fieldResultsInFile,
  fieldResultsShown,
  foldedCaptureNodes,
  inlineTokens,
  listNumberFieldDocx,
  liveFoldFaults,
  openDocx,
  type ParagraphSpec,
  positionOfText,
} from "./__tests__/listNumberFieldFixture";
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
      return state.apply(state.tr.replaceRange(at, at, unfoldPastedListNumberFields(copied)));
    }
    default: {
      const unhandled: never = step;
      throw new Error(`Unhandled step ${JSON.stringify(unhandled)}`);
    }
  }
};

const captureMarkup = (doc: PMNode): string[] =>
  foldedCaptureNodes(doc).map(({ node }) => String(node.attrs["xml"]));

describe("the LISTNUM fields of a document under splits, joins, typing and pastes", () => {
  test(
    "no field is hidden unless a marker shows it, and none is hidden twice or lost",
    async () => {
      const parsed = await openDocx(await listNumberFieldDocx(SPECS));
      const opened = editorState(parsed);
      const original = captureMarkup(opened.doc);
      // The field behind "one " and the field behind "two " are on the line.
      expect(original).toHaveLength(4);
      expect(liveFoldFaults(opened.doc)).toEqual([]);

      assertProperty(
        fc.property(fc.array(stepArbitrary, { minLength: 1, maxLength: 12 }), (steps) => {
          let state = opened;
          for (const step of steps) {
            state = applyStep(state, step);
            expect(liveFoldFaults(state.doc)).toEqual([]);
            // Two tabs read from the same markup are alike, so captures are
            // counted, not told apart: none may stand more often than it did.
            const left = [...original];
            for (const xml of captureMarkup(state.doc)) {
              const at = left.indexOf(xml);
              expect(at).not.toBe(-1);
              left.splice(at, 1);
            }
          }

          const paragraphs = bodyParagraphs(fromProseDoc(state.doc, parsed));
          const written: string[] = [];
          for (const paragraph of paragraphs) {
            const tokens = inlineTokens(serializeParagraph(paragraph));
            expect(fieldResultsShown(paragraph)).toBe(fieldResultsInFile(tokens));
            written.push(...tokens);
          }
          // Nothing here deletes, so every field is still written at least once.
          for (const result of RESULTS) {
            expect(written).toContain(result);
          }
        }),
        {
          numRuns: 150,
          // Text typed at the very start of a paragraph whose marker shows a field.
          examples: [
            [[{ kind: "type", at: { paragraph: 0, offset: 0 } }]],
            // A whole paragraph copied and pasted ahead of itself.
            [
              [
                {
                  kind: "paste",
                  from: { paragraph: 0, offset: 0 },
                  to: { paragraph: 0.25, offset: 0 },
                  at: { paragraph: 0, offset: 0 },
                },
              ],
            ],
          ],
        },
      );
    },
    propertyTestTimeout(120_000),
  );
});

/**
 * The pass runs after every transaction of every document, so what it costs a
 * document it has nothing to do in is the cost that counts. It is stated here
 * as the paragraphs it looks at, which no machine's speed changes.
 */
describe("the paragraphs the editor's pass looks at", () => {
  const PARAGRAPHS = 2000;

  const plainParagraphs = (): PMNode[] =>
    Array.from({ length: PARAGRAPHS }, (_, index) =>
      schema.node("paragraph", null, [schema.text(`Clause ${index} of the agreement`)]),
    );

  /** One word replaced in every other plain paragraph, back to front, in one transaction. */
  const replaceInEveryOther = (state: EditorState): Transaction => {
    const starts: number[] = [];
    let plain = 0;
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
    state.doc.forEach((node, position) => {
      if (!node.textContent.startsWith("Clause")) {
        return;
      }
      if (plain % 2 === 0) {
        starts.push(position);
      }
      plain += 1;
    });
    const tr = state.tr;
    for (const start of starts.toReversed()) {
      tr.insertText("Section", start + 1, start + 1 + "Clause".length);
    }
    return tr;
  };

  const counted = (doc: PMNode): { state: EditorState; visits: () => number } => {
    let visits = 0;
    const plugin = foldedListNumberPlugin({
      onParagraphVisit: () => {
        visits += 1;
      },
    });
    return { state: EditorState.create({ doc, plugins: [plugin] }), visits: () => visits };
  };

  test("none, for a thousand steps over a document that holds no field", () => {
    const { state, visits } = counted(schema.node("doc", null, plainParagraphs()));
    const tr = replaceInEveryOther(state);
    expect(tr.steps).toHaveLength(PARAGRAPHS / 2);

    const next = state.apply(tr);

    expect(next.doc.firstChild?.textContent.startsWith("Section")).toBe(true);
    expect(visits()).toBe(0);
  });

  test("a handful, when one paragraph among them holds a field", async () => {
    const paraId = SPECS[1]?.paraId ?? "";
    const parsed = await openDocx(await listNumberFieldDocx(SPECS.slice(1, 2)));
    const folded = toProseDoc(parsed).firstChild;
    if (!folded || foldedCaptureNodes(folded).length === 0) {
      throw new Error("The fixture opens with a paragraph whose marker hides a field");
    }
    const { state, visits } = counted(schema.node("doc", null, [folded, ...plainParagraphs()]));

    // A thousand steps, none of them in the paragraph that holds the field.
    state.apply(replaceInEveryOther(state));
    expect(visits()).toBe(0);

    // The same thousand, and one more in that paragraph.
    const tr = replaceInEveryOther(state);
    const edited = state.apply(
      tr.insertText("!", positionOfText(state.doc, paraId, "second body") + 1),
    );

    expect(tr.steps).toHaveLength(PARAGRAPHS / 2 + 1);
    expect(visits()).toBeGreaterThanOrEqual(1);
    expect(visits()).toBeLessThanOrEqual(3);
    expect(liveFoldFaults(edited.doc)).toEqual([]);
  });
});
