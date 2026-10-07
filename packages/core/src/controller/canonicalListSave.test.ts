import { expect, test } from "bun:test";
import { Fragment, Slice } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { createEmptyDocument } from "../utils/createDocument";
import { schema } from "../prosemirror/schema";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";
import { prepareCanonicalPaste } from "./canonicalClipboard";
import { validateDocxPackage } from "@stll/docx-core";
import { FOLIO_DOCX_SERIALIZATION_MODE } from "../types/docxSerialization";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { mintListInstance } from "../docx/listNumberingInstances";
import { createNumberingMap } from "../docx/numberingParser";
import { listItemAttrs } from "../prosemirror/listNumbering";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import { CANONICAL_SAVE_FALLBACK_DIAGNOSTIC } from "../../../../test/canonicalSaveDiagnostics";

test(
  "generated pasted lists preserve save/reopen fidelity through full-save fallbacks and history",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(fc.constantFrom("Alpha", "é", "😀", "東京"), { minLength: 2, maxLength: 4 }),
        fc.integer({ min: 0, max: 2 }),
        async (texts, nestedLevel) => {
          const source = await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" }));
          const input = (await prepareCanonicalDocxInput(source)).unwrap();
          const session = createCanonicalSession(
            await parseDocx(input, { preloadFonts: false }),
          ).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 6)));
          const minted = mintListInstance(undefined, { kind: "bullet" });
          const numbering = createNumberingMap(minted.definitions);
          const attrs = expectParagraphAttrs(schema.node("paragraph"));
          const slice = new Slice(
            Fragment.fromArray(
              texts.map((text, index) =>
                schema.node(
                  "paragraph",
                  listItemAttrs(
                    attrs,
                    { numId: minted.numId, ilvl: index === 1 ? nestedLevel : 0 },
                    numbering,
                  ),
                  schema.text(text),
                ),
              ),
            ),
            0,
            0,
          );
          state = publishCanonicalProjection({
            session,
            state,
            commit: prepareCanonicalPaste({ session, state, slice }).unwrap(),
          }).unwrap().state;
          const pasted = session.document;
          for (const transition of ["pasted", "undo", "redo"] as const) {
            if (transition !== "pasted") {
              state = publishCanonicalProjection({
                session,
                state,
                commit: (transition === "undo"
                  ? session.prepareUndo(state)
                  : session.prepareRedo(state)
                ).unwrap(),
              }).unwrap().state;
            }
            for (const mode of Object.values(FOLIO_DOCX_SERIALIZATION_MODE)) {
              const snapshot = session.captureSaveSnapshot();
              const saved = await serializeCanonicalSave({ snapshot, options: { mode } });
              expect(saved.diagnostics).toEqual(
                mode === FOLIO_DOCX_SERIALIZATION_MODE.preferSelective && transition !== "undo"
                  ? [CANONICAL_SAVE_FALLBACK_DIAGNOSTIC]
                  : [],
              );
              expect(await validateDocxPackage(new Uint8Array(saved.buffer))).toEqual({
                valid: true,
              });
              const reopened = await parseDocx(saved.buffer, { preloadFonts: false });
              expect(structuredClone(reopened.package.document.content)).toEqual(
                structuredClone(snapshot.document.package.document.content),
              );
              expect(reopened.package.numbering).toEqual(snapshot.document.package.numbering);
            }
          }
          expect(session.document).toEqual(pasted);
        },
      ),
      { numRuns: 5 },
    );
  },
  propertyTestTimeout(30_000),
);
