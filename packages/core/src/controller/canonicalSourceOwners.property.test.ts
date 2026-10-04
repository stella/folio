import {
  cloneDocumentWithParagraphPropertySources,
  getDocumentParagraphPropertySourceContract,
  getParagraphPropertySource,
  getParagraphPropertySourceToken,
  paragraphPropertySourceTokenMatchesContract,
  paragraphPropertySourceBelongsToDocument,
} from "../docx/paragraphPropertySource";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { schema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";
import { canonicalJson } from "../utils/canonicalJson";
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

test(
  "cross-paragraph history preserves exact owners while reopening establishes fresh captures",
  async () => {
    let cases = 0;
    await assertProperty(
      fc.asyncProperty(text, text, async (initial, replacement) => {
        // Exercise both property-source classes on every generated case.
        for (const sourceKind of ["captured", "uncaptured"] as const) {
          cases += 1;
          const document = createEmptyDocument({ initialText: `a${initial}` });
          const first = document.package.document.content.at(0);
          if (first?.type !== "paragraph") throw new TypeError("Expected paragraph seed");
          first.paraId = "F0220001";
          first.formatting = { styleId: "Normal" };
          document.package.document.content.push({
            type: "paragraph",
            paraId: "F0220002",
            content: [],
            ...(sourceKind === "captured" ? { formatting: { styleId: "Normal" } } : {}),
          });
          const parsed = await parseDocx(await createDocx(document), {
            preloadFonts: false,
            detectVariables: false,
          });
          const session = createCanonicalSession(parsed).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          // Keep the prefix so its authored properties cross onto the trailing survivor.
          const projectedFirst = state.doc.firstChild;
          if (projectedFirst === null) throw new TypeError("Expected projected paragraph");
          const from = projectedFirst.nodeSize - 1;
          const to = state.doc.content.size - 1;
          state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
          const before = { document: session.document, projection: state.doc.toJSON() };
          state = publishCanonicalProjection({
            state,
            session,
            commit: session.prepareReplace(state, { from, to, text: replacement }).unwrap(),
          }).unwrap().state;
          const after = { document: session.document, projection: state.doc.toJSON() };
          const checkCaptures = (snapshot: typeof before) => {
            for (const original of snapshot.document.package.document.content) {
              if (original.type !== "paragraph") throw new TypeError("Expected snapshot paragraph");
              const current = session.document.package.document.content.find(
                (block) => block.type === "paragraph" && block.paraId === original.paraId,
              );
              if (current?.type !== "paragraph") throw new TypeError("Lost snapshot paragraph");
              expect(getParagraphPropertySource(current)).toBe(
                getParagraphPropertySource(original),
              );
            }
          };
          const survivor = session.document.package.document.content.at(0);
          if (survivor?.type !== "paragraph") throw new TypeError("Expected joined survivor");
          expect(getParagraphPropertySource(survivor) !== undefined).toBe(
            sourceKind === "captured",
          );
          for (let cycle = 0; cycle < 2; cycle++) {
            state = publishCanonicalProjection({
              state,
              session,
              commit: session.prepareUndo(state).unwrap(),
            }).unwrap().state;
            expect(session.document).toEqual(before.document);
            expect(state.doc.toJSON()).toEqual(before.projection);
            checkCaptures(before);
            state = publishCanonicalProjection({
              state,
              session,
              commit: session.prepareRedo(state).unwrap(),
            }).unwrap().state;
            expect(session.document).toEqual(after.document);
            expect(state.doc.toJSON()).toEqual(after.projection);
            checkCaptures(after);
          }
          const reopened = await parseDocx(await createDocx(session.document), {
            preloadFonts: false,
            detectVariables: false,
          });
          expect(canonicalJson(reopened.package.document.content)).toEqual(
            canonicalJson(session.document.package.document.content),
          );
          const contract = getDocumentParagraphPropertySourceContract(reopened);
          if (contract === undefined) throw new TypeError("Missing reopened source contract");
          for (const paragraph of reopened.package.document.content) {
            if (paragraph.type !== "paragraph") throw new TypeError("Expected reopened paragraph");
            expect(
              paragraphPropertySourceTokenMatchesContract(
                getParagraphPropertySourceToken(paragraph),
                contract,
              ),
            ).toBe(true);
            expect(getParagraphPropertySource(paragraph)).toBeDefined();
            expect(paragraphPropertySourceBelongsToDocument(paragraph, reopened)).toBe(true);
          }
        }
      }),
      { numRuns: 24 },
    );
    expect(cases).toBeGreaterThan(0);
  },
  propertyTestTimeout(30_000),
);
