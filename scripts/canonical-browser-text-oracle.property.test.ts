import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import { schema } from "../packages/core/src/prosemirror/schema";
import { canonicalTextSelection } from "../tests/parity/canonicalTextSelection";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../packages/core/src/controller/canonicalSession";
import { createEmptyDocument } from "../packages/core/src/utils/createDocument";

setDefaultTimeout(propertyTestTimeout(5_000));

const text = fc
  .array(fc.constantFrom("a", "é", "😀", "東京", "é"), { maxLength: 5 })
  .map((parts) => parts.join(""));

test("document-position text slices agree with flat text replacement across paragraphs", () => {
  assertProperty(
    fc.property(
      fc.array(text, { minLength: 1, maxLength: 5 }),
      fc.nat(),
      fc.nat(),
      fc.nat(),
      text,
      (paragraphs, ordinal, anchor, head, inserted) => {
        const index = ordinal % paragraphs.length;
        const target = paragraphs.at(index);
        if (target === undefined) throw new TypeError("Missing generated paragraph");
        const content = paragraphs.map((value) =>
          schema.nodes["paragraph"].create(null, value === "" ? [] : schema.text(value)),
        );
        const doc = schema.nodes["doc"].create(null, content);
        let start = 1;
        for (let prior = 0; prior < index; prior++) start += doc.child(prior).nodeSize;
        const gaps = [0];
        for (const character of target) gaps.push((gaps.at(-1) ?? 0) + character.length);
        const left = gaps.at(anchor % gaps.length);
        const right = gaps.at(head % gaps.length);
        if (left === undefined || right === undefined)
          throw new TypeError("Missing generated text boundary");
        const from = Math.min(left, right);
        const to = Math.max(left, right);
        const state = EditorState.create({
          schema,
          doc,
          selection: TextSelection.create(doc, start + from, start + to),
        });
        const slices = canonicalTextSelection(state);
        expect(slices.before + slices.selected + slices.after).toBe(paragraphs.join(""));
        const expected = [...paragraphs];
        expected[index] = target.slice(0, from) + inserted + target.slice(to);
        expect(slices.before + inserted + slices.after).toBe(expected.join(""));
        expect(
          state.apply(state.tr.insertText(inserted, state.selection.from, state.selection.to)).doc
            .textContent,
        ).toBe(expected.join(""));
      },
    ),
    { numRuns: 100 },
  );
});

test("canonical split followed by typing uses text at the caret's paragraph", () => {
  const source = createEmptyDocument({ initialText: "alpha😀café東京" });
  const first = source.package.document.content.at(0);
  if (first?.type !== "paragraph") throw new TypeError("Missing paragraph fixture");
  first.paraId = "F0220001";
  const session = createCanonicalSession(source).unwrap();
  let state = EditorState.create({
    schema,
    doc: session.projection.doc,
    selection: TextSelection.create(session.projection.doc, 1, 6),
  });
  state = publishCanonicalProjection({
    state,
    session,
    commit: session.prepareSplit(state).unwrap(),
  }).unwrap().state;
  const slices = canonicalTextSelection(state);
  const before = state.doc.textContent;
  const from = state.selection.from;
  const to = state.selection.to;
  state = publishCanonicalProjection({
    state,
    session,
    commit: session.prepareReplace(state, { from, to, text: "alpha" }).unwrap(),
  }).unwrap().state;
  expect(state.doc.textContent).toBe("alpha😀café東京");
  expect(slices.before + "alpha" + slices.after).toBe(state.doc.textContent);
  expect(before.slice(0, from - 1) + "alpha" + before.slice(to - 1)).not.toBe(
    state.doc.textContent,
  );
});
