import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { schema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(30_000));

const text = fc
  .array(fc.constantFrom("a", "é", "😀", "東京", "é"), { maxLength: 5 })
  .map((parts) => parts.join(""));

test("repeated undo restores every removed paragraph's original source ownership", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.array(text, { minLength: 2, maxLength: 5 }),
      fc.nat(),
      text,
      async (paragraphs, ordinal, replacement) => {
        const document = createEmptyDocument({ initialText: "" });
        document.package.document.content = paragraphs.map((value, index) => ({
          type: "paragraph",
          paraId: (index + 1).toString(16).padStart(8, "0").toUpperCase(),
          content: value === "" ? [] : [{ type: "run", content: [{ type: "text", text: value }] }],
        }));
        const parsed = await parseDocx(await createDocx(document), {
          preloadFonts: false,
          detectVariables: false,
        });
        const session = createCanonicalSession(parsed).unwrap();
        let state = EditorState.create({ schema, doc: session.projection.doc });
        const start = ordinal % (paragraphs.length - 1);
        let from = 1;
        for (let index = 0; index < start; index++) from += state.doc.child(index).nodeSize;
        const to = state.doc.content.size - 1;
        state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
        const before = state.doc.toJSON();
        const original = cloneDocumentWithParagraphPropertySources(session.document);
        state = publishCanonicalProjection({
          state,
          session,
          commit: session.prepareReplace(state, { from, to, text: replacement }).unwrap(),
        }).unwrap().state;
        const after = state.doc.toJSON();
        for (let cycle = 0; cycle < 2; cycle++) {
          state = publishCanonicalProjection({
            state,
            session,
            commit: session.prepareUndo(state).unwrap(),
          }).unwrap().state;
          expect(session.document).toEqual(original);
          expect(state.doc.toJSON()).toEqual(before);
          state = publishCanonicalProjection({
            state,
            session,
            commit: session.prepareRedo(state).unwrap(),
          }).unwrap().state;
          expect(state.doc.toJSON()).toEqual(after);
        }
      },
    ),
    { numRuns: 24 },
  );
});
