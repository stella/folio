import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import type { Mark, Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { assertProperty, propertyTestTimeout } from "../../../../../../test/property-testing";
import { getCanonicalCommandIntents } from "../../canonicalCommands";
import { schema, singletonManager } from "../../schema";
import { HyperlinkRemovalRefusal, removeHyperlinkInRange } from "../../hyperlinkRemoval";
import { createSuggestionModePlugin } from "../../plugins/suggestionMode";
import { acceptAllChanges, rejectAllChanges } from "../../commands/comments";

setDefaultTimeout(propertyTestTimeout(15_000));

const segments = fc.array(
  fc.record({
    length: fc.integer({ min: 1, max: 5 }),
    bold: fc.boolean(),
    italic: fc.boolean(),
  }),
  { minLength: 2, maxLength: 6 },
);
const neighboringAttributes = fc.constantFrom(
  { href: "https://other.example/" },
  { tooltip: "Other" },
  { rId: "rId2" },
  { _docxHyperlinkIndex: 2 },
);

type AssertRemovalOptions = {
  nodes: readonly PMNode[];
  link: Mark;
  neighbor: Mark;
};

const assertRemoval = ({ nodes, link, neighbor }: AssertRemovalOptions) => {
  const length = nodes.reduce((total, node) => total + node.nodeSize, 0);
  const from = 2;
  const to = from + length;
  const doc = schema.node("doc", undefined, [
    schema.node("paragraph", { paraId: "12345678" }, [
      schema.text("L", [neighbor]),
      ...nodes,
      schema.text("R", [neighbor]),
      schema.text(" "),
      schema.text("island", [link]),
    ]),
  ]);
  const expected = doc.type.create(doc.attrs, [
    schema.node("paragraph", { paraId: "12345678" }, [
      schema.text("L", [neighbor]),
      ...nodes.map((node) => node.mark(link.removeFromSet(node.marks))),
      schema.text("R", [neighbor]),
      schema.text(" "),
      schema.text("island", [link]),
    ]),
  ]);
  const command = singletonManager.requireCommand("removeHyperlink")();
  for (let cursor = from + 1; cursor < to; cursor += 1) {
    let state = EditorState.create({
      schema,
      doc,
      selection: TextSelection.create(doc, cursor),
    });
    if (doc.resolve(cursor).parent.type.name !== "paragraph") {
      expect(() => command(state)).toThrow(HyperlinkRemovalRefusal);
      expect(() => getCanonicalCommandIntents(command, state)).toThrow(HyperlinkRemovalRefusal);
      expect(() =>
        command(state, () => {
          throw new TypeError("A refused command dispatched");
        }),
      ).toThrow(HyperlinkRemovalRefusal);
      expect(state.doc.eq(doc)).toBe(true);
      continue;
    }
    const intents = getCanonicalCommandIntents(command, state);
    expect(command(state)).toBe(true);
    expect(
      command(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);
    expect(state.doc.eq(expected)).toBe(true);
    expect(state.doc.textContent).toBe(doc.textContent);
    expect(intents).toEqual([{ type: "removeHyperlink", from, to }]);
  }
};

test("collapsed removal covers exactly one equal hyperlink across formatting segments", () => {
  assertProperty(
    fc.property(segments, neighboringAttributes, (pieces, different) => {
      const attrs = {
        href: "https://example.com/",
        tooltip: "Original",
        rId: "rId1",
        _docxHyperlinkIndex: 1,
      };
      const link = schema.mark("hyperlink", attrs);
      const neighbor = schema.mark("hyperlink", { ...attrs, ...different });
      const nodes = pieces.map(({ length, bold, italic }, index) =>
        schema.text(String.fromCharCode(97 + index).repeat(length), [
          link,
          ...(bold ? [schema.mark("bold")] : []),
          ...(italic ? [schema.mark("italic")] : []),
          // Force a run boundary even when the generated formatting agrees.
          ...(index % 2 === 0 ? [schema.mark("underline")] : []),
        ]),
      );
      assertRemoval({ nodes, link, neighbor });
    }),
    {
      numRuns: 60,
      id: "collapsed removal covers exactly one equal hyperlink across formatting segments",
    },
  );
});

const atomFactories = {
  image: (link: Mark) =>
    schema.node(
      "image",
      { src: "https://example.com/image.png", width: 16, height: 16 },
      undefined,
      [link],
    ),
  field: (link: Mark) =>
    schema.node("field", { fieldType: "PAGE", instruction: "PAGE", displayText: "1" }, undefined, [
      link,
    ]),
  structuredField: (link: Mark) =>
    schema.node(
      "structuredField",
      { fieldType: "REF", instruction: "REF target", displayText: "result" },
      [schema.text("result")],
      [link],
    ),
  tab: (link: Mark) => schema.node("tab", undefined, undefined, [link]),
} as const;

test("collapsed removal crosses linked inline atoms without losing content", () => {
  assertProperty(
    fc.property(segments, neighboringAttributes, (pieces, different) => {
      const attrs = {
        href: "https://example.com/",
        tooltip: "Original",
        rId: "rId1",
        _docxHyperlinkIndex: 1,
      };
      const link = schema.mark("hyperlink", attrs);
      const neighbor = schema.mark("hyperlink", { ...attrs, ...different });
      const exercised = new Set<string>();
      for (const [kind, createAtom] of Object.entries(atomFactories)) {
        const nodes: PMNode[] = [];
        for (const [index, { length, bold, italic }] of pieces.entries()) {
          if (index > 0) nodes.push(createAtom(link));
          nodes.push(
            schema.text("x".repeat(length), [
              link,
              ...(bold ? [schema.mark("bold")] : []),
              ...(italic ? [schema.mark("italic")] : []),
            ]),
          );
        }
        assertRemoval({ nodes, link, neighbor });
        exercised.add(kind);
      }
      expect([...exercised].sort()).toEqual(Object.keys(atomFactories).sort());
    }),
    { numRuns: 40, id: "collapsed removal crosses linked inline atoms without losing content" },
  );
});

test("collapsed removal unlinks text/image/text as one hyperlink", () => {
  const link = schema.mark("hyperlink", { href: "https://example.com/" });
  const neighbor = schema.mark("hyperlink", { href: "https://other.example/" });
  assertRemoval({
    nodes: [
      schema.text("before", [link]),
      atomFactories.image(link),
      schema.text("after", [link, schema.mark("bold")]),
    ],
    link,
    neighbor,
  });
});

const assertPartialFieldRefusal = (text: string) => {
  const link = schema.mark("hyperlink", { href: "https://example.com/" });
  const field = schema.node(
    "structuredField",
    { fieldType: "REF", instruction: "REF target", displayText: text },
    [schema.text(text)],
    [link],
  );
  const doc = schema.node("doc", undefined, [
    schema.node("paragraph", { paraId: "12345678" }, [
      schema.text("L", [link]),
      field,
      schema.text("R", [link]),
    ]),
  ]);
  const command = singletonManager.requireCommand("removeHyperlink")();
  // Field starts at 2; its content spans [3, 3 + text.length).
  for (let from = 3; from < 3 + text.length; from += 1) {
    for (let to = from + 1; to <= 3 + text.length; to += 1) {
      const state = EditorState.create({
        schema,
        doc,
        selection: TextSelection.create(doc, from, to),
        plugins: [createSuggestionModePlugin(true, "Reviewer")],
      });
      const tr = state.tr;
      expect(() => removeHyperlinkInRange(state, tr, from, to)).toThrow(HyperlinkRemovalRefusal);
      expect(tr.steps).toEqual([]);
      expect(tr.doc.eq(doc)).toBe(true);
      expect(() => command(state)).toThrow(HyperlinkRemovalRefusal);
      expect(() =>
        command(state, () => {
          throw new TypeError("A refused command dispatched");
        }),
      ).toThrow(HyperlinkRemovalRefusal);
      for (const resolve of [acceptAllChanges(), rejectAllChanges()]) {
        let resolved = state;
        resolve(state, (transaction) => {
          resolved = state.apply(transaction);
        });
        expect(resolved.doc.eq(doc)).toBe(true);
        expect(resolved.doc.textContent).toBe(`L${text}R`);
      }
    }
  }
};

test("suggesting removal refuses a linked structured-field substring before replacement", () => {
  assertPartialFieldRefusal("result");
});

test("every generated field substring refuses atom splitting with exact accept and reject content", () => {
  assertProperty(
    fc.property(fc.integer({ min: 2, max: 8 }), (length) =>
      assertPartialFieldRefusal("x".repeat(length)),
    ),
    {
      numRuns: 20,
      id: "every generated field substring refuses atom splitting with exact accept and reject content",
    },
  );
});
