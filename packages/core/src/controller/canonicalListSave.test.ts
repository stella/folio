import { expect, test } from "bun:test";
import { Fragment, Slice } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Paragraph } from "../types/document";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { createEmptyDocument } from "../utils/createDocument";
import { schema } from "../prosemirror/schema";
import {
  createCanonicalSession,
  CanonicalSessionError,
  publishCanonicalProjection,
} from "./canonicalSession";
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
import { normalizeCanonicalListRendering } from "./canonicalListRendering";
import { CANONICAL_SAVE_FALLBACK_DIAGNOSTIC } from "../../../../test/canonicalSaveDiagnostics";

test(
  "generated LISTNUM level edits preserve child counters, history and save/reopen",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 1, max: 6 }),
        fc.integer({ min: 1, max: 3 }),
        fc.constantFrom("leading", "inside"),
        async (level, fieldCount, position) => {
          const minted = mintListInstance(undefined, { kind: "numbered" });
          for (const abstract of minted.definitions.abstractNums) {
            for (const entry of abstract.levels) {
              entry.numFmt = "decimal";
              entry.lvlText = `%${entry.ilvl + 1}.`;
              entry.start = 1;
            }
          }
          const text = {
            type: "run",
            content: [{ type: "text", text: "Body" }],
          } satisfies Paragraph["content"][number];
          const fields = Array.from(
            { length: fieldCount },
            () =>
              ({
                type: "complexField",
                fieldType: "LISTNUM",
                instruction: "LISTNUM",
                fieldCode: [],
                fieldResult: [{ type: "run", content: [{ type: "text", text: "1" }] }],
              }) satisfies Paragraph["content"][number],
          );
          const host = {
            type: "paragraph",
            paraId: "10000001",
            formatting: { numPr: { kind: "reference", numId: minted.numId, ilvl: 0 } },
            content: position === "leading" ? [...fields, text] : [text, ...fields],
          } satisfies Paragraph;
          const following = {
            type: "paragraph",
            paraId: "10000002",
            formatting: { numPr: { kind: "reference", numId: minted.numId, ilvl: level + 1 } },
            content: [text],
          } satisfies Paragraph;
          const parsed = await parseDocx(
            await createDocx({
              package: { numbering: minted.definitions, document: { content: [host, following] } },
            }),
            { preloadFonts: false },
          );
          const session = createCanonicalSession(parsed).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          const before = structuredClone(session.document);
          const at = session.projection.inputAddressAt(1).unwrap();
          state = publishCanonicalProjection({
            session,
            state,
            commit: session
              .prepareIntent(state, {
                type: "formatParagraph",
                at,
                patch: { numPr: { kind: "reference", numId: minted.numId, ilvl: level } },
              })
              .unwrap(),
          }).unwrap().state;
          const after = structuredClone(session.document);
          const marker = (document: typeof parsed) => {
            const paragraph = document.package.document.content.at(1);
            if (paragraph?.type !== "paragraph") throw new TypeError("Missing child paragraph");
            return paragraph.listRendering?.marker;
          };
          expect(marker(session.document)).toBe(`${fieldCount + 1}.`);
          for (const mode of Object.values(FOLIO_DOCX_SERIALIZATION_MODE)) {
            const saved = await serializeCanonicalSave({
              snapshot: session.captureSaveSnapshot(),
              options: { mode },
            });
            const reopened = await parseDocx(saved.buffer, { preloadFonts: false });
            expect(marker(reopened)).toBe(`${fieldCount + 1}.`);
            expect(structuredClone(reopened.package.document.content)).toEqual(
              after.package.document.content,
            );
          }
          state = publishCanonicalProjection({
            session,
            state,
            commit: session.prepareUndo(state).unwrap(),
          }).unwrap().state;
          expect(structuredClone(session.document)).toEqual(before);
          state = publishCanonicalProjection({
            session,
            state,
            commit: session.prepareRedo(state).unwrap(),
          }).unwrap().state;
          expect(structuredClone(session.document)).toEqual(after);
          expect(marker(session.document)).toBe(`${fieldCount + 1}.`);
        },
      ),
      {
        examples: [
          [1, 1, "inside"],
          [1, 1, "leading"],
        ],
      },
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

const listEdits = [
  "remove",
  "outdent",
  "level",
  "levelOnly",
  "idOnly",
  "split",
  "joinBackward",
  "joinForward",
] as const;

const listEditParagraph = {
  remove: 1,
  outdent: 1,
  level: 1,
  levelOnly: 1,
  idOnly: 1,
  split: 1,
  joinBackward: 3,
  joinForward: 2,
} as const satisfies Record<(typeof listEdits)[number], number>;

test(
  "generated list edits normalize rendering and preserve exact history and save fixed points",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom("bullet", "numbered"),
        fc.integer({ min: 1, max: 7 }),
        fc.constantFrom("Alpha", "é", "😀", "東京"),
        fc.constantFrom("direct", "style"),
        fc.integer({ min: 0, max: 12 }),
        async (kind, level, text, provenance, start) => {
          const minted = mintListInstance(undefined, { kind, start });
          const source = await createDocx({
            package: {
              numbering: minted.definitions,
              ...(provenance === "direct"
                ? {}
                : {
                    styles: {
                      styles: [
                        {
                          type: "paragraph",
                          styleId: "List",
                          name: "List",
                          pPr: { numPr: { kind: "reference", numId: minted.numId, ilvl: level } },
                        },
                      ],
                    },
                  }),
              document: {
                content: [0, level, 0, undefined].map((ilvl, index) => {
                  const paragraph = {
                    type: "paragraph",
                    paraId: (index + 1).toString(16).padStart(8, "0"),
                    content: [{ type: "run", content: [{ type: "text", text }] }],
                  } satisfies Paragraph;
                  if (ilvl === undefined) return paragraph;
                  const formatting = (
                    provenance === "direct"
                      ? { numPr: { kind: "reference", numId: minted.numId, ilvl } }
                      : { styleId: "List", numPr: { kind: "levelOnly", ilvl } }
                  ) satisfies NonNullable<Paragraph["formatting"]>;
                  return Object.assign(paragraph, { formatting });
                }),
              },
            },
          });
          const parsed = await parseDocx(source, { preloadFonts: false });
          const original = structuredClone(parsed);
          const exercised = new Set<string>();
          for (const edit of listEdits) {
            const session = createCanonicalSession(parsed).unwrap();
            let state = EditorState.create({ schema, doc: session.projection.doc });
            const index = listEditParagraph[edit];
            let position = 1;
            for (let previous = 0; previous < index; previous++)
              position += state.doc.child(previous).nodeSize;
            if (edit === "joinForward") position += state.doc.child(index).content.size;
            state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, position)));
            const before = session.document;
            const beforeProjection = state.doc.toJSON();
            const beforeSelection = state.selection.toJSON();
            const at = session.projection.inputAddressAt(position).unwrap();
            const commit = (() => {
              switch (edit) {
                case "remove":
                  return session.prepareIntent(state, {
                    type: "formatParagraph",
                    at,
                    patch: { numPr: { kind: "none" } },
                  });
                case "outdent":
                case "level":
                  return session.prepareIntent(state, {
                    type: "formatParagraph",
                    at,
                    patch: {
                      numPr: {
                        kind: "reference",
                        numId: minted.numId,
                        ilvl: edit === "outdent" ? level - 1 : level + 1,
                      },
                    },
                  });
                case "levelOnly":
                  return session.prepareIntent(state, {
                    type: "formatParagraph",
                    at,
                    patch: {
                      numPr: { kind: "levelOnly", ilvl: level + 1 },
                    },
                  });
                case "idOnly":
                  return session.prepareIntent(state, {
                    type: "formatParagraph",
                    at,
                    patch: {
                      numPr: { kind: "reference", numId: minted.numId },
                    },
                  });
                case "split":
                  return session.prepareSplit(state);
                case "joinBackward":
                  return session.prepareJoin(state, "backward");
                case "joinForward":
                  return session.prepareJoin(state, "forward");
                default: {
                  const unreachable: never = edit;
                  return unreachable;
                }
              }
            })().unwrap();
            state = publishCanonicalProjection({ session, state, commit }).unwrap().state;
            const after = session.document;
            const afterProjection = state.doc.toJSON();
            const afterSelection = state.selection.toJSON();
            expect(normalizeCanonicalListRendering(after).document).toBe(after);
            expect(structuredClone(parsed)).toEqual(original);
            for (const transition of ["edited", "undo", "redo"] as const) {
              if (transition !== "edited") {
                state = publishCanonicalProjection({
                  session,
                  state,
                  commit: (transition === "undo"
                    ? session.prepareUndo(state)
                    : session.prepareRedo(state)
                  ).unwrap(),
                }).unwrap().state;
                expect(session.document).toEqual(transition === "undo" ? before : after);
                expect(state.doc.toJSON()).toEqual(
                  transition === "undo" ? beforeProjection : afterProjection,
                );
                expect(state.selection.toJSON()).toEqual(
                  transition === "undo" ? beforeSelection : afterSelection,
                );
              }
              for (const mode of Object.values(FOLIO_DOCX_SERIALIZATION_MODE)) {
                const snapshot = session.captureSaveSnapshot();
                const saved = await serializeCanonicalSave({ snapshot, options: { mode } });
                const reopened = await parseDocx(saved.buffer, { preloadFonts: false });
                expect(structuredClone(reopened.package.document.content)).toEqual(
                  structuredClone(snapshot.document.package.document.content),
                );
                expect(reopened.package.numbering).toEqual(snapshot.document.package.numbering);
              }
            }
            exercised.add(edit);
          }
          expect([...exercised].toSorted()).toEqual([...listEdits].toSorted());
        },
      ),
      { numRuns: 5 },
    );
  },
  propertyTestTimeout(30_000),
);

test(
  "activation validates identities before normalizing inherited list overrides",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(fc.integer({ min: 0, max: 2 }), { minLength: 1, maxLength: 4 }),
        fc.constantFrom("bullet" as const, "numbered" as const),
        async (levels, kind) => {
          const minted = mintListInstance(undefined, { kind });
          const document = createEmptyDocument();
          document.package.numbering = minted.definitions;
          document.package.document.content = levels.map((ilvl) => ({
            type: "paragraph",
            formatting: {
              numPrFromStyle: { kind: "reference", numId: minted.numId, ilvl: 0 },
              numPr: { kind: "levelOnly", ilvl },
            },
            content: [{ type: "run", content: [{ type: "text", text: "List item" }] }],
          }));
          const original = structuredClone(document);
          const refused = createCanonicalSession(document);
          expect(refused.isErr()).toBe(true);
          if (refused.isErr()) expect(refused.error).toBeInstanceOf(CanonicalSessionError);
          expect(document).toStrictEqual(original);
          const input = (await prepareCanonicalDocxInput(await createDocx(document))).unwrap();
          const identified = await parseDocx(input, { preloadFonts: false });
          // Restore the generated style inheritance/override shape after the load-time ID owner.
          for (const [index, block] of identified.package.document.content.entries()) {
            if (block.type !== "paragraph") continue;
            const ilvl = levels.at(index);
            if (ilvl === undefined) continue;
            block.formatting = {
              numPrFromStyle: { kind: "reference", numId: minted.numId, ilvl: 0 },
              numPr: { kind: "levelOnly", ilvl },
            };
          }
          const identifiedOriginal = structuredClone(identified);
          const duplicate = structuredClone(identified);
          const first = duplicate.package.document.content.at(0);
          expect(first).toBeDefined();
          if (first === undefined) return;
          duplicate.package.document.content.push(structuredClone(first));
          const duplicateOriginal = structuredClone(duplicate);
          const duplicateResult = createCanonicalSession(duplicate);
          expect(duplicateResult.isErr()).toBe(true);
          if (duplicateResult.isErr())
            expect(duplicateResult.error).toBeInstanceOf(CanonicalSessionError);
          expect(duplicate).toStrictEqual(duplicateOriginal);
          const result = createCanonicalSession(identified);
          expect(result.isOk()).toBe(true);
          const session = result.unwrap();
          const identities = new Set<string>();
          for (const [index, block] of session.document.package.document.content.entries()) {
            expect(block.type).toBe("paragraph");
            if (block.type !== "paragraph") continue;
            expect(block.paraId).toBeDefined();
            if (block.paraId !== undefined) identities.add(block.paraId);
            expect(block.formatting?.numPr).toEqual({
              kind: "reference",
              numId: minted.numId,
              ilvl: levels.at(index),
            });
            expect(block.listRendering?.level).toBe(levels.at(index));
          }
          expect(identities.size).toBe(levels.length);
          expect(structuredClone(identified)).toStrictEqual(identifiedOriginal);
          expect(document).toStrictEqual(original);
        },
      ),
      { numRuns: 20 },
    );
  },
  propertyTestTimeout(5_000),
);
