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
import JSZip from "jszip";
import {
  findWordprocessingChild,
  findChildrenByNamespaceUri,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "../docx/xmlParser";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { mintListInstance } from "../docx/listNumberingInstances";
import { createNumberingMap } from "../docx/numberingParser";
import { listItemAttrs } from "../prosemirror/listNumbering";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import { CANONICAL_SAVE_FALLBACK_DIAGNOSTIC } from "../../../../test/canonicalSaveDiagnostics";

test(
  "generated numbering sources survive edits, save/reopen and later style reference changes",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 0, max: 4 }),
        fc.integer({ min: 0, max: 4 }),
        fc.constantFrom("bullet" as const, "numbered" as const),
        async (styleLevel, directLevel, kind) => {
          const original = mintListInstance(undefined, { kind });
          const replacement = mintListInstance(original.definitions, { kind });
          const authored = [
            undefined,
            { kind: "levelOnly", ilvl: directLevel },
            { kind: "reference", numId: original.numId, ilvl: styleLevel },
            { kind: "reference", numId: original.numId, ilvl: directLevel },
          ] as const;
          const parsed = await parseDocx(
            await createDocx({
              package: {
                numbering: replacement.definitions,
                styles: {
                  styles: [
                    {
                      type: "paragraph",
                      styleId: "Numbered",
                      name: "Numbered",
                      pPr: {
                        numPr: { kind: "reference", numId: original.numId, ilvl: styleLevel },
                      },
                    },
                  ],
                },
                document: {
                  content: authored.map((numPr, index) => ({
                    type: "paragraph",
                    paraId: (index + 1).toString(16).padStart(8, "0"),
                    formatting: { styleId: "Numbered", ...(numPr === undefined ? {} : { numPr }) },
                    content: [{ type: "run", content: [{ type: "text", text: "Item" }] }],
                  })),
                },
              },
            }),
            { preloadFonts: false },
          );
          const session = createCanonicalSession(parsed).unwrap();
          const state = EditorState.create({ schema, doc: session.projection.doc });
          publishCanonicalProjection({
            session,
            state,
            commit: session.prepareReplace(state, { from: 1, to: 1, text: "X" }).unwrap(),
          }).unwrap();
          for (const mode of Object.values(FOLIO_DOCX_SERIALIZATION_MODE)) {
            const saved = await serializeCanonicalSave({
              snapshot: session.captureSaveSnapshot(),
              options: { mode },
            });
            const xml = await (
              await JSZip.loadAsync(saved.buffer)
            )
              .file("word/document.xml")
              ?.async("text");
            expect(xml).toBeDefined();
            if (xml === undefined) return;
            const paragraphs = findChildrenByNamespaceUri(
              findWordprocessingChild(parseXmlDocument(xml), "body"),
              WORDPROCESSINGML_NAMESPACE_URIS,
              "p",
            );
            expect(paragraphs).toHaveLength(authored.length);
            const reopened = await parseDocx(saved.buffer, { preloadFonts: false });
            for (const [index, paragraph] of reopened.package.document.content.entries()) {
              expect(paragraph.type).toBe("paragraph");
              if (paragraph.type !== "paragraph") continue;
              const expected = authored.at(index);
              expect(paragraph.formatting?.numPr).toEqual(expected);
              const numbering = findWordprocessingChild(
                findWordprocessingChild(paragraphs.at(index), "pPr"),
                "numPr",
              );
              expect(findWordprocessingChild(numbering, "numId") !== null).toBe(
                expected?.kind === "reference",
              );
              expect(findWordprocessingChild(numbering, "ilvl") !== null).toBe(
                expected !== undefined,
              );
            }
            // A new package changes the stylesheet itself; existing paragraph overrides
            // must survive that change rather than following an equal former value.
            const restyledPackage = structuredClone(reopened.package);
            const style = restyledPackage.styles?.styles.find(
              ({ styleId }) => styleId === "Numbered",
            );
            expect(style).toBeDefined();
            if (style === undefined) return;
            style.pPr = {
              numPr: { kind: "reference", numId: replacement.numId, ilvl: styleLevel },
            };
            const restyled = await parseDocx(await createDocx({ package: restyledPackage }), {
              preloadFonts: false,
            });
            for (const [index, paragraph] of restyled.package.document.content.entries()) {
              if (paragraph.type !== "paragraph") continue;
              const expected = authored.at(index);
              expect(paragraph.formatting?.numPr).toEqual(expected);
              expect(paragraph.listRendering?.numId).toBe(
                expected?.kind === "reference" ? original.numId : replacement.numId,
              );
              expect(paragraph.listRendering?.level).toBe(
                expected === undefined ? styleLevel : expected.ilvl,
              );
            }
          }
        },
      ),
      { numRuns: 5 },
    );
  },
  propertyTestTimeout(30_000),
);

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
              const xml = await (
                await JSZip.loadAsync(saved.buffer)
              )
                .file("word/document.xml")
                ?.async("text");
              if (xml === undefined) throw new Error("Saved package must contain its document XML");
              const paragraphs = findChildrenByNamespaceUri(
                findWordprocessingChild(parseXmlDocument(xml), "body"),
                WORDPROCESSINGML_NAMESPACE_URIS,
                "p",
              );
              expect(paragraphs.length).toBeGreaterThan(0);
              // These lists author numbering only; a save must not materialize level indentation.
              for (const paragraph of paragraphs) {
                expect(
                  findWordprocessingChild(findWordprocessingChild(paragraph, "pPr"), "ind"),
                ).toBeNull();
              }
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
