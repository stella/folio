import { expect, test } from "bun:test";
import { panic } from "better-result";
import { storyRevisionIds } from "../../../../test/reviewProjection";
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
  type CanonicalSession,
  CanonicalSessionError,
  publishCanonicalProjection,
} from "./canonicalSession";
import { prepareCanonicalPaste } from "./canonicalClipboard";
import { validateDocxPackage } from "@stll/docx-core";
import { DOCUMENT_OP_TYPES, OP_STORIES, type DocumentOp } from "@stll/docx-core/ops";
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
        id: "generated LISTNUM level edits preserve child counters, history and save/reopen",
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
          for (const [index, paragraph] of session.document.package.document.content.entries()) {
            if (paragraph.type !== "paragraph") continue;
            expect(paragraph.formatting?.numPr).toEqual(authored.at(index));
          }
          expect(normalizeCanonicalListRendering(session.document).document).toBe(session.document);
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
      {
        id: "generated numbering sources survive edits, save/reopen and later style reference changes",
        numRuns: 5,
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
      {
        numRuns: 5,
        id: "generated pasted lists preserve save/reopen fidelity through full-save fallbacks and history",
      },
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
        fc.constantFrom("inherited", "zero", "matching", "custom"),
        async (kind, level, text, provenance, start, indentationProfile) => {
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
                  const indentation = (() => {
                    switch (indentationProfile) {
                      case "inherited":
                        return {};
                      case "zero":
                        return { indentLeft: 0, indentFirstLine: 0, hangingIndent: false };
                      case "matching":
                        return (
                          minted.definitions.abstractNums
                            .at(0)
                            ?.levels.find((entry) => entry.ilvl === ilvl)?.pPr ??
                          panic("A minted list level must have indentation.")
                        );
                      case "custom":
                        return {
                          indentLeft: 360,
                          indentRight: 120,
                          indentFirstLine: 120,
                          hangingIndent: false,
                        };
                      default: {
                        const unreachable: never = indentationProfile;
                        return unreachable;
                      }
                    }
                  })();
                  return Object.assign(paragraph, {
                    formatting: { ...formatting, ...indentation },
                  });
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
            const prepareEdit = (editSession: CanonicalSession, editState: EditorState) => {
              const at = editSession.projection.inputAddressAt(position).unwrap();
              switch (edit) {
                case "remove":
                  return editSession.prepareIntent(editState, {
                    type: "formatParagraph",
                    at,
                    patch: { numPr: { kind: "none" } },
                  });
                case "outdent":
                case "level":
                  return editSession.prepareIntent(editState, {
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
                  return editSession.prepareIntent(editState, {
                    type: "formatParagraph",
                    at,
                    patch: {
                      numPr: { kind: "levelOnly", ilvl: level + 1 },
                    },
                  });
                case "idOnly":
                  return editSession.prepareIntent(editState, {
                    type: "formatParagraph",
                    at,
                    patch: {
                      numPr: { kind: "reference", numId: minted.numId },
                    },
                  });
                case "split":
                  return editSession.prepareSplit(editState);
                case "joinBackward":
                  return editSession.prepareJoin(editState, "backward");
                case "joinForward":
                  return editSession.prepareJoin(editState, "forward");
                default: {
                  const unreachable: never = edit;
                  return unreachable;
                }
              }
            };
            const commit = prepareEdit(session, state).unwrap();
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
            const suggested = createCanonicalSession(parsed).unwrap();
            suggested.setMode({ type: "suggesting", author: "Reviewer" });
            let suggestedState = EditorState.create({ schema, doc: suggested.projection.doc });
            suggestedState = suggestedState.apply(
              suggestedState.tr.setSelection(TextSelection.create(suggestedState.doc, position)),
            );
            const beforeSuggestion = structuredClone(suggested.document);
            suggestedState = publishCanonicalProjection({
              session: suggested,
              state: suggestedState,
              commit: prepareEdit(suggested, suggestedState).unwrap(),
            }).unwrap().state;
            // Rejection must restore authored properties, even when the snapshot
            // began with derived list indentation and is loaded in a fresh session.
            const reviewing = createCanonicalSession(suggested.document).unwrap();
            let reviewState = EditorState.create({ schema, doc: reviewing.projection.doc });
            reviewState = publishCanonicalProjection({
              session: reviewing,
              state: reviewState,
              commit: reviewing
                .prepareResolve(reviewState, {
                  revisionIds: storyRevisionIds(reviewing.document),
                  resolution: "reject",
                })
                .unwrap(),
            }).unwrap().state;
            expect(structuredClone(reviewing.document)).toEqual(beforeSuggestion);
            expect(reviewState.doc.toJSON()).toEqual(beforeProjection);
            exercised.add(edit);
          }
          expect([...exercised].toSorted()).toEqual([...listEdits].toSorted());
        },
      ),
      {
        numRuns: 5,
        id: "generated list edits normalize rendering and preserve exact history and save fixed points",
      },
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
          const state = EditorState.create({ schema, doc: session.projection.doc });
          const anchor = session.document.package.document.content.at(0);
          if (anchor?.type !== "paragraph" || anchor.paraId === undefined)
            return panic("Missing identified fixture paragraph.");
          const before = structuredClone(session.document);
          for (const identity of [undefined, anchor.paraId]) {
            const inserted: Paragraph = {
              type: "paragraph",
              ...(identity === undefined ? {} : { paraId: identity }),
              formatting: {
                numPrFromStyle: { kind: "reference", numId: minted.numId, ilvl: 0 },
                numPr: { kind: "levelOnly", ilvl: levels.at(0) ?? 0 },
              },
              content: [{ type: "run", content: [{ type: "text", text: "Inserted list item" }] }],
            };
            const operation = {
              type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
              story: OP_STORIES.MAIN,
              at: { type: "before", blockId: anchor.paraId },
              blocks: [inserted],
            } as const satisfies DocumentOp;
            const refusedInsertion = session.prepareOperations(state, [operation]);
            expect(refusedInsertion.isErr()).toBe(true);
            if (refusedInsertion.isErr())
              expect(refusedInsertion.error).toBeInstanceOf(CanonicalSessionError);
            expect(structuredClone(session.document)).toStrictEqual(before);
          }
          expect(identities.size).toBe(levels.length);
          expect(structuredClone(identified)).toStrictEqual(identifiedOriginal);
          expect(document).toStrictEqual(original);
        },
      ),
      {
        numRuns: 20,
        id: "activation validates identities before normalizing inherited list overrides",
      },
    );
  },
  propertyTestTimeout(5_000),
);
