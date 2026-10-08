import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";
import { assertProperty, propertyTestTimeout } from "../../../../../../test/property-testing";
import { getCanonicalCommandIntents } from "../../canonicalCommands";
import { schema, singletonManager } from "../../schema";

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
    }),
    { numRuns: 60 },
  );
});
