import { expect, test } from "bun:test";
import fc from "fast-check";
import { validateDocxPackage } from "../packages/docx-core/src/validate/docx";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../packages/core/src/controller/canonicalSession";
import { schema } from "../packages/core/src/prosemirror/schema";
import { toProseDoc } from "../packages/core/src/prosemirror/conversion/toProseDoc";
import { createEmptyDocument } from "../packages/core/src/utils/createDocument";
import { createDocx } from "../packages/core/src/docx/rezip";
import { parseDocx } from "../packages/core/src/docx/parser";
import {
  PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR,
  PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR,
} from "../packages/core/src/docx/paragraphPropertySource";
import { createMissingOpBurndown } from "./canonical-missing-ops";
import { assertProperty, propertyTestTimeout } from "./property-testing";

// Source ownership is package-specific; reopening intentionally establishes a new contract.
const portableProjection = (state: EditorState["doc"]) =>
  JSON.stringify(state.toJSON(), (key, value) =>
    key === PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR || key === PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR
      ? undefined
      : value,
  );

const textArbitrary = fc
  .array(fc.constantFrom("a", "é", "😀", "東京", "é"), {
    minLength: 1,
    maxLength: 6,
  })
  .map((characters) => characters.join(""));
const actionArbitrary = fc.record({
  kind: fc.constantFrom("replace", "delete", "cross-paragraph", "structural-input"),
  anchor: fc.nat(100),
  head: fc.nat(100),
  text: textArbitrary,
});

test(
  "canonical property flows preserve projection, exact history and saved content",
  async () => {
    const missing = createMissingOpBurndown();
    await assertProperty(
      fc.asyncProperty(
        textArbitrary,
        fc.array(actionArbitrary, { minLength: 1, maxLength: 24 }),
        async (initial, actions) => {
          const generated = createEmptyDocument({ initialText: initial });
          const first = generated.package.document.content.at(0);
          if (first?.type !== "paragraph") throw new TypeError("Expected paragraph seed");
          first.paraId = "F0220001";
          generated.package.document.content.push({
            type: "paragraph",
            paraId: "F0220002",
            content: [],
          });
          const source = await parseDocx(await createDocx(generated), {
            preloadFonts: false,
            detectVariables: false,
          });
          const session = createCanonicalSession(source).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          const checkProjection = () => {
            expect(state.doc.eq(session.projection.doc)).toBe(true);
            expect(
              state.doc.eq(
                toProseDoc(
                  session.document,
                  session.document.package.styles
                    ? { styles: session.document.package.styles }
                    : undefined,
                ),
              ),
            ).toBe(true);
          };
          for (const action of actions) {
            const paragraph = state.doc.firstChild;
            if (!paragraph) throw new TypeError("Lost generated paragraph");
            const gaps = [1];
            let cursor = 1;
            for (const character of paragraph.textContent) {
              cursor += character.length;
              gaps.push(cursor);
            }
            const anchor = gaps.at(action.anchor % gaps.length);
            const head = gaps.at(action.head % gaps.length);
            if (anchor === undefined || head === undefined)
              throw new TypeError("Lost generated boundary");
            const from = Math.min(anchor, head);
            const to =
              action.kind === "cross-paragraph"
                ? state.doc.content.size - 1
                : Math.max(anchor, head);
            state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
            const before = {
              document: session.document,
              projection: state.doc.toJSON(),
              selection: state.selection.toJSON(),
              version: session.version,
              undo: session.canUndo,
              redo: session.canRedo,
            };
            let text = action.text;
            if (action.kind === "structural-input") text = "\n";
            if (action.kind === "delete") text = "";
            const input = { from, to, text };
            const prepared = session.prepareReplace(state, input);
            if (prepared.isErr()) {
              missing.record(action.kind);
              expect(session.document).toEqual(before.document);
              expect(session.version).toBe(before.version);
              expect(session.canUndo).toBe(before.undo);
              expect(session.canRedo).toBe(before.redo);
              checkProjection();
              continue;
            }
            state = publishCanonicalProjection({ state, session, commit: prepared.value }).unwrap()
              .state;
            checkProjection();
            const expectedText =
              paragraph.textContent.slice(0, from - 1) + text + paragraph.textContent.slice(to - 1);
            expect(state.doc.firstChild?.textContent).toBe(expectedText);
            const after = {
              document: session.document,
              projection: state.doc.toJSON(),
              selection: state.selection.toJSON(),
            };
            state = publishCanonicalProjection({
              state,
              session,
              commit: session.prepareUndo(state).unwrap(),
            }).unwrap().state;
            checkProjection();
            expect(session.document).toEqual(before.document);
            expect(state.doc.toJSON()).toEqual(before.projection);
            expect(state.selection.toJSON()).toEqual(before.selection);
            state = publishCanonicalProjection({
              state,
              session,
              commit: session.prepareRedo(state).unwrap(),
            }).unwrap().state;
            checkProjection();
            expect(session.document).toEqual(after.document);
            expect(state.doc.toJSON()).toEqual(after.projection);
            expect(state.selection.toJSON()).toEqual(after.selection);
            const saved = await createDocx(session.document);
            expect(await validateDocxPackage(saved)).toEqual({ valid: true });
            const reopened = await parseDocx(saved, {
              preloadFonts: false,
              detectVariables: false,
            });
            expect(reopened.package.document.content).toEqual(
              session.document.package.document.content,
            );
            const reopenedSession = createCanonicalSession(reopened).unwrap();
            expect(portableProjection(reopenedSession.projection.doc)).toEqual(
              portableProjection(state.doc),
            );
          }
        },
      ),
      { numRuns: 20 },
    );
    console.log(missing.markdown());
  },
  propertyTestTimeout(120_000),
);
