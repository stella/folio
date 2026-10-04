import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { panic } from "better-result";
import { EditorState, TextSelection } from "prosemirror-state";
import { paragraphLogicalText, type TextPosition } from "@stll/docx-core/ops";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document } from "../types/document";
import { schema, singletonManager } from "../prosemirror/schema";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import { directionToAuthoredBidi } from "../prosemirror/paragraphDirection";
import { TAB_STOP_ALIGNMENT_VALUES, TAB_LEADER_VALUES } from "../types/documentEnumValues";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import {
  assignDocumentParagraphPropertySourceContract,
  cloneDocumentWithParagraphPropertySources,
  assignParagraphPropertySource,
  getParagraphPropertySource,
  getParagraphPropertySourceToken,
} from "../docx/paragraphPropertySource";
import {
  createCanonicalSession,
  publishCanonicalProjection,
  type CanonicalCommit,
  type CanonicalSession,
} from "./canonicalSession";
import { prepareCanonicalAutoformat, prepareCanonicalCommands } from "./canonicalStructure";

setDefaultTimeout(propertyTestTimeout(240_000));

const seed = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          content: [{ type: "run", content: [{ type: "text", text: "ab😀cd" }] }],
        },
        {
          type: "paragraph",
          paraId: "00000001",
          content: [{ type: "run", content: [{ type: "text", text: "EF" }] }],
        },
      ],
    },
  },
});

const paragraphs = (session: CanonicalSession) =>
  session.document.package.document.content.map((paragraph) => {
    if (paragraph.type !== "paragraph" || paragraph.paraId === undefined)
      return panic("Canonical test lost a paragraph identity.");
    return paragraph;
  });
const texts = (session: CanonicalSession) => paragraphs(session).map(paragraphLogicalText);
const accept = (session: CanonicalSession, state: EditorState, commit: CanonicalCommit) => {
  const applied = publishCanonicalProjection({ session, state, commit }).unwrap();
  expect(applied.state.doc.eq(session.projection.doc)).toBe(true);
  const identities = paragraphs(session).map(({ paraId }) => paraId);
  expect(new Set(identities).size).toBe(identities.length);
  const sourceTokens = paragraphs(session).flatMap((paragraph) => {
    const token = getParagraphPropertySourceToken(paragraph);
    return token === undefined ? [] : [token];
  });
  expect(new Set(sourceTokens).size).toBe(sourceTokens.length);
  for (const paragraph of paragraphs(session)) {
    const address =
      session.projection.paragraph(paragraph.paraId ?? "") ?? panic("Projection lost a paragraph.");
    expect(address.node.textBetween(0, address.node.content.size, "", "\uFFFC")).toBe(
      paragraphLogicalText(paragraph),
    );
  }
  return applied.state;
};
const select = (state: EditorState, from: number, to = from) =>
  state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));

const kinds = ["split", "join", "run", "paragraph", "atom", "list", "input"] as const;
const inputArbitrary = fc.record({
  kind: fc.constantFrom(...kinds),
  offset: fc.nat(1000),
  reverse: fc.boolean(),
  bold: fc.boolean(),
});

test("list-level commands preserve plain paragraphs and refuse all-plain selections without history", async () => {
  let cases = 0;
  await assertProperty(
    fc.property(
      fc.record({
        indent: fc.integer({ min: 1, max: 4000 }),
        level: fc.integer({ min: 1, max: 7 }),
        reverse: fc.boolean(),
      }),
      ({ indent: generatedIndent, level, reverse }) => {
        cases += 1;
        for (const indent of [undefined, 0, generatedIndent]) {
          for (const direction of ["increase", "decrease"] as const) {
            for (const kind of ["plain", "mixed"] as const) {
              const document = seed();
              for (const paragraph of document.package.document.content) {
                if (paragraph.type !== "paragraph") panic("Expected paragraph fixture.");
                paragraph.formatting = indent === undefined ? {} : { indentLeft: indent };
              }
              if (kind === "mixed") {
                const builder = createCanonicalSession(document).unwrap();
                let state = EditorState.create({ schema, doc: builder.projection.doc });
                const last = builder.projection.doc.child(0).nodeSize + 1;
                state = select(state, last);
                accept(
                  builder,
                  state,
                  prepareCanonicalCommands(builder, state, [
                    { type: "toggleList", kind: "decimal" },
                  ]).unwrap(),
                );
                Object.assign(document, builder.document);
                const list = document.package.document.content.at(1);
                if (list?.type !== "paragraph" || list.formatting?.numPr?.kind !== "reference")
                  panic("Expected list fixture.");
                list.formatting.numPr.ilvl = level;
              }
              const session = createCanonicalSession(document).unwrap();
              let state = EditorState.create({ schema, doc: session.projection.doc });
              const end = state.doc.content.size - 1;
              state = reverse ? select(state, end, 1) : select(state, 1, end);
              const before = session.document;
              const selection = state.selection.toJSON();
              const prepared = prepareCanonicalCommands(session, state, [
                { type: "changeListLevel", direction },
              ]);
              if (kind === "plain") {
                expect(prepared.isErr()).toBe(true);
                if (prepared.isErr()) expect(prepared.error.reason).toBe("noChange");
                expect(session.document).toStrictEqual(before);
                expect(session.version).toBe(0);
                expect(session.canUndo).toBe(false);
                continue;
              }
              state = accept(session, state, prepared.unwrap());
              expect(session.document.package.document.content.at(0)).toStrictEqual(
                before.package.document.content.at(0),
              );
              const list = paragraphs(session).at(1);
              expect(
                list?.formatting?.numPr?.kind === "reference"
                  ? list.formatting.numPr.ilvl
                  : undefined,
              ).toBe(level + (direction === "increase" ? 1 : -1));
              state = accept(session, state, session.prepareUndo(state).unwrap());
              expect(session.document).toStrictEqual(before);
              expect(state.selection.toJSON()).toEqual(selection);
              expect(session.canUndo).toBe(false);
            }
          }
        }
      },
    ),
  );
  expect(cases).toBeGreaterThan(0);
});

const autoformatMarkerArbitrary = fc.oneof(
  fc.constantFrom("*", "-").map((marker) => ({ type: "bullet", marker }) as const),
  fc
    .record({ start: fc.integer({ min: 1, max: 999 }), punctuation: fc.constantFrom(".", ")") })
    .map(
      ({ start, punctuation }) =>
        ({ type: "numbered", marker: `${start}${punctuation}`, start, punctuation }) as const,
    ),
  fc
    .integer({ min: 1, max: 6 })
    .map((level) => ({ type: "heading", marker: "#".repeat(level), level }) as const),
);

describe("canonical structural commands", () => {
  test("generated structural sequences preserve text, identities, projection and exact history selections", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(inputArbitrary, { minLength: 14, maxLength: 28 }),
        async (generated) => {
          const document = seed();
          assignDocumentParagraphPropertySourceContract(document, "b".repeat(64));
          const session = createCanonicalSession(document).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          let oracle = ["ab😀cd", "EF"];
          const history: {
            before: Document;
            after: Document;
            preSelection: ReturnType<typeof state.selection.toJSON>;
            postSelection: ReturnType<typeof state.selection.toJSON>;
          }[] = [];
          const outcomes = { applied: 0, noChange: 0 };
          const exercised = new Set<(typeof kinds)[number]>();
          const repeatedSplits = [0, 1].map(
            (offset) => ({ kind: "split", offset, reverse: false, bold: false }) as const,
          );
          const inputs = [
            ...kinds.map((kind) => ({ kind, offset: 2, reverse: false, bold: true })),
            ...repeatedSplits,
            ...generated,
          ];
          for (const input of inputs) {
            exercised.add(input.kind);
            const sources = paragraphs(session);
            const index =
              input.kind === "run"
                ? sources.findIndex((paragraph) => paragraphLogicalText(paragraph).length > 0)
                : input.offset % sources.length;
            const source = sources.at(index) ?? panic("Sequence has no addressed paragraph.");
            const address =
              session.projection.paragraph(source.paraId ?? "") ??
              panic("Sequence lost its address.");
            const gaps = [0];
            for (const character of address.text) gaps.push((gaps.at(-1) ?? 0) + character.length);
            const offset =
              gaps.at(input.offset % gaps.length) ?? panic("Sequence lost a code-point gap.");
            state = select(state, address.start + offset);
            let prepare: ReturnType<typeof session.prepareIntent>;
            let expectedCaret: TextPosition | undefined;
            const before = session.document;
            const previousVersion = session.version;
            const previousUndo = session.canUndo;
            const previousRedo = session.canRedo;
            let preSelection = state.selection.toJSON();
            switch (input.kind) {
              case "split": {
                prepare = session.prepareSplit(state);
                expectedCaret = { story: "main", blockId: source.paraId ?? "", offset: 0 };
                const original = oracle.at(index) ?? panic("Text oracle lost a paragraph.");
                oracle.splice(index, 1, original.slice(0, offset), original.slice(offset));
                break;
              }
              case "join": {
                if (sources.length === 1) {
                  prepare = session.prepareJoin(select(state, address.start), "backward");
                  break;
                }
                const firstIndex = input.offset % (sources.length - 1);
                const first = sources.at(firstIndex) ?? panic("Missing join paragraph.");
                const firstAddress =
                  session.projection.paragraph(first.paraId ?? "") ??
                  panic("Missing join address.");
                state = select(state, firstAddress.start + firstAddress.text.length);
                preSelection = state.selection.toJSON();
                prepare = session.prepareJoin(state, "forward");
                expectedCaret = {
                  story: "main",
                  blockId: sources.at(firstIndex + 1)?.paraId ?? "",
                  offset: firstAddress.text.length,
                };
                oracle.splice(
                  firstIndex,
                  2,
                  (oracle.at(firstIndex) ?? "") + (oracle.at(firstIndex + 1) ?? ""),
                );
                break;
              }
              case "run": {
                const from = address.start;
                const to = from + address.text.length;
                state = select(state, input.reverse ? to : from, input.reverse ? from : to);
                preSelection = state.selection.toJSON();
                prepare = prepareCanonicalCommands(session, state, [
                  { type: "formatRun", from, to, patch: { bold: input.bold } },
                ]);
                break;
              }
              case "paragraph":
                prepare = prepareCanonicalCommands(session, state, [
                  {
                    type: "formatParagraph",
                    at: state.selection.head,
                    patch: { alignment: input.reverse ? "right" : "center" },
                  },
                ]);
                break;
              case "atom": {
                const at = session.projection.addressAt(state.selection.head).unwrap();
                expectedCaret = { ...at, offset: at.offset + 1 };
                const atom = (() => {
                  if (input.offset % 3 === 0) return { type: "break", breakType: "page" } as const;
                  if (input.reverse) return { type: "break", breakType: "textWrapping" } as const;
                  return { type: "tab" } as const;
                })();
                prepare = session.prepareIntent(state, {
                  type: "insertAtom",
                  from: at,
                  to: at,
                  atom,
                });
                const original = oracle.at(index) ?? panic("Text oracle lost an atom paragraph.");
                oracle[index] = original.slice(0, offset) + "\uFFFC" + original.slice(offset);
                break;
              }
              case "list":
                prepare = prepareCanonicalCommands(session, state, [
                  { type: "toggleList", kind: input.reverse ? "bullet" : "decimal" },
                ]);
                break;
              case "input": {
                expectedCaret = { story: "main", blockId: source.paraId ?? "", offset: offset + 1 };
                prepare = session.prepareReplace(state, {
                  from: state.selection.head,
                  to: state.selection.head,
                  text: "é",
                  semantic: "replacement",
                });
                const original = oracle.at(index) ?? panic("Text oracle lost an input paragraph.");
                oracle[index] = original.slice(0, offset) + "é" + original.slice(offset);
                break;
              }
              default: {
                const unreachable: never = input.kind;
                return panic(`Unknown sequence kind ${unreachable}`);
              }
            }
            if (prepare.isErr()) {
              expect(prepare.error.reason).toBe("noChange");
              expect(["run", "paragraph", "join"]).toContain(input.kind);
              if (input.kind === "join") expect(sources.length).toBe(1);
              outcomes.noChange += 1;
              expect(session.document).toStrictEqual(before);
              expect(session.version).toBe(previousVersion);
              expect(session.canUndo).toBe(previousUndo);
              expect(session.canRedo).toBe(previousRedo);
              expect(state.selection.toJSON()).toEqual(preSelection);
              expect(texts(session)).toEqual(oracle);
              continue;
            }
            outcomes.applied += 1;
            state = accept(session, state, prepare.value);
            expect(texts(session)).toEqual(oracle);
            if (expectedCaret !== undefined) {
              const selection = session.projection.selectionAt(state).unwrap();
              expect(selection.anchor.blockId).toBe(expectedCaret.blockId);
              expect(selection.anchor.offset).toBe(expectedCaret.offset);
              expect(selection.head).toEqual(selection.anchor);
            }
            if (input.kind === "run" || input.kind === "paragraph" || input.kind === "list")
              expect(state.selection.toJSON()).toEqual(preSelection);
            if (input.kind === "run") {
              const updated =
                paragraphs(session).at(index) ?? panic("Formatted paragraph disappeared.");
              expect(
                updated.content.every(
                  (run) => run.type === "run" && run.formatting?.bold === input.bold,
                ),
              ).toBe(true);
            }
            if (input.kind === "paragraph")
              expect(paragraphs(session).at(index)?.formatting?.alignment).toBe(
                input.reverse ? "right" : "center",
              );
            history.push({
              before,
              after: session.document,
              preSelection,
              postSelection: state.selection.toJSON(),
            });
          }
          expect(outcomes.applied + outcomes.noChange).toBe(inputs.length);
          expect(outcomes.applied).toBe(history.length);
          expect(outcomes.applied).toBeGreaterThanOrEqual(kinds.length);
          expect([...exercised].sort()).toEqual([...kinds].sort());
          for (const entry of history.toReversed()) {
            state = accept(session, state, session.prepareUndo(state).unwrap());
            expect(session.document).toStrictEqual(entry.before);
            expect(state.selection.toJSON()).toEqual(entry.preSelection);
          }
          expect(session.canUndo).toBe(false);
          for (const entry of history) {
            state = accept(session, state, session.prepareRedo(state).unwrap());
            expect(session.document).toStrictEqual(entry.after);
            expect(state.selection.toJSON()).toEqual(entry.postSelection);
          }
          expect(session.canRedo).toBe(false);
        },
      ),
      { numRuns: 24 },
    );
  });

  test("marker rules atomically remove their marker, apply paragraph meaning and restore it on undo", async () => {
    await assertProperty(
      fc.asyncProperty(
        autoformatMarkerArbitrary,
        fc.boolean(),
        async (generated, styleAvailable) => {
          const { marker } = generated;
          const document = seed();
          document.package.document.content = [
            {
              type: "paragraph",
              paraId: "12345678",
              content: [{ type: "run", content: [{ type: "text", text: marker }] }],
            },
          ];
          if (generated.type === "heading" && styleAvailable) {
            document.package.styles = {
              styles: [
                {
                  styleId: `Heading${generated.level}`,
                  type: "paragraph",
                  name: `Heading ${generated.level}`,
                },
              ],
            };
          }
          const session = createCanonicalSession(document).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          state = select(state, marker.length + 1);
          const before = session.document;
          const selection = state.selection.toJSON();
          const prepared = prepareCanonicalAutoformat(session, state, {
            from: state.selection.head,
            to: state.selection.head,
            text: " ",
          });
          if (generated.type === "heading" && !styleAvailable) {
            expect(prepared).toBeUndefined();
            expect(session.document).toStrictEqual(before);
            expect(state.selection.toJSON()).toEqual(selection);
            return;
          }
          const commit = prepared ?? panic("Declared marker was not recognized.");
          state = accept(session, state, commit.unwrap());
          expect(texts(session)).toEqual([""]);
          expect(state.selection.head).toBe(1);
          const paragraph = paragraphs(session).at(0) ?? panic("Rule paragraph disappeared.");
          if (generated.type === "heading")
            expect(paragraph.formatting?.styleId).toBe(`Heading${generated.level}`);
          else {
            expect(paragraph.formatting?.numPr?.kind).toBe("reference");
            expect(session.document.package.numbering?.nums.length).toBe(1);
            expect(state.doc.firstChild?.attrs["listNumFmt"]).toBe(
              generated.type === "bullet" ? "bullet" : "decimal",
            );
            if (generated.type === "numbered") {
              const levels = session.document.package.numbering?.abstractNums.at(0)?.levels;
              expect(levels?.map(({ start }) => start)).toEqual([
                generated.start,
                ...Array.from({ length: 8 }, () => 1),
              ]);
              expect(levels?.map(({ lvlText }) => lvlText)).toEqual(
                Array.from({ length: 9 }, (_, index) => `%${index + 1}${generated.punctuation}`),
              );
            }
          }
          state = accept(session, state, session.prepareUndo(state).unwrap());
          expect(session.document).toEqual(before);
          expect(state.selection.toJSON()).toEqual(selection);
          expect(session.canUndo).toBe(false);
          state = accept(session, state, session.prepareRedo(state).unwrap());
          expect(texts(session)).toEqual([""]);
        },
      ),
      { numRuns: 30 },
    );
  });

  test("numbered autoformat keeps its start at level zero through every nested level", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.record({
          start: fc.integer({ min: 2, max: 999 }),
          punctuation: fc.constantFrom(".", ")"),
        }),
        async ({ start, punctuation }) => {
          const document = seed();
          const marker = `${start}${punctuation}`;
          document.package.document.content = [
            {
              type: "paragraph",
              paraId: "12345678",
              content: [{ type: "run", content: [{ type: "text", text: marker }] }],
            },
          ];
          const session = createCanonicalSession(document).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          state = select(state, marker.length + 1);
          const history: {
            before: Document;
            after: Document;
            preSelection: ReturnType<typeof state.selection.toJSON>;
            postSelection: ReturnType<typeof state.selection.toJSON>;
          }[] = [];
          const apply = (commit: CanonicalCommit) => {
            const before = cloneDocumentWithParagraphPropertySources(session.document);
            const preSelection = state.selection.toJSON();
            state = accept(session, state, commit);
            history.push({
              before,
              after: cloneDocumentWithParagraphPropertySources(session.document),
              preSelection,
              postSelection: state.selection.toJSON(),
            });
          };
          const initial = cloneDocumentWithParagraphPropertySources(session.document);
          const initialSelection = state.selection.toJSON();
          const prepared = prepareCanonicalAutoformat(session, state, {
            from: state.selection.head,
            to: state.selection.head,
            text: " ",
          });
          apply(prepared?.unwrap() ?? panic("Numbered marker was not recognized."));

          const expectedStarts = [start, ...Array.from({ length: 8 }, () => 1)];
          const expectedFormats = Array.from({ length: 9 }, () => "decimal");
          for (let level = 0; level <= 8; level += 1) {
            if (level > 0) {
              apply(
                prepareCanonicalCommands(session, state, [
                  { type: "changeListLevel", direction: "increase" },
                ]).unwrap(),
              );
            }
            const paragraph = paragraphs(session).at(0) ?? panic("List paragraph disappeared.");
            const numbering = paragraph.formatting?.numPr;
            expect(numbering?.kind === "reference" ? numbering.ilvl : undefined).toBe(level);
            const attrs =
              state.doc.firstChild?.attrs ?? panic("Rendered list paragraph disappeared.");
            expect(attrs["listLevelStarts"]).toEqual(expectedStarts.slice(0, level + 1));
            expect(attrs["listLevelNumFmts"]).toEqual(expectedFormats.slice(0, level + 1));
            expect(attrs["listMarkerTemplate"]).toBe(`%${level + 1}${punctuation}`);
          }

          for (const entry of history.toReversed()) {
            state = accept(session, state, session.prepareUndo(state).unwrap());
            expect(session.document).toStrictEqual(entry.before);
            expect(state.selection.toJSON()).toStrictEqual(entry.preSelection);
          }
          expect(session.document).toStrictEqual(initial);
          expect(state.selection.toJSON()).toStrictEqual(initialSelection);
          expect(session.canUndo).toBe(false);
          for (const entry of history) {
            state = accept(session, state, session.prepareRedo(state).unwrap());
            expect(session.document).toStrictEqual(entry.after);
            expect(state.selection.toJSON()).toStrictEqual(entry.postSelection);
          }
        },
      ),
      { numRuns: 16 },
    );
  });

  test("note-reference digits and inline atoms do not become autoformat markers", async () => {
    await assertProperty(
      fc.asyncProperty(fc.constantFrom(".", ")"), async (punctuation) => {
        for (const type of ["footnoteRef", "endnoteRef"] as const) {
          for (const id of [1, 12]) {
            for (const atom of ["none", "tab", "break"] as const) {
              const document = seed();
              document.package.document.content = [
                {
                  type: "paragraph",
                  paraId: "12345678",
                  content: [
                    {
                      type: "run",
                      content: [
                        { type, id },
                        ...(atom === "tab" ? [{ type: "tab" as const }] : []),
                        ...(atom === "break"
                          ? [{ type: "break" as const, breakType: "page" as const }]
                          : []),
                        { type: "text", text: punctuation },
                      ],
                    },
                  ],
                },
              ];
              const session = createCanonicalSession(document).unwrap();
              let state = EditorState.create({ schema, doc: session.projection.doc });
              const address =
                session.projection.paragraph("12345678") ?? panic("Note paragraph is absent.");
              const logicalEnd = session.projection
                .positionAt({
                  story: session.projection.story,
                  blockId: address.blockId,
                  offset: address.text.length,
                })
                .unwrap();
              expect(session.projection.addressAt(logicalEnd).unwrap().offset).toBe(
                address.text.length,
              );
              expect(logicalEnd).toBe(address.start + address.node.content.size);
              state = select(state, logicalEnd);
              const before = structuredClone(session.document);
              const selection = state.selection.toJSON();
              const prepared = prepareCanonicalAutoformat(session, state, {
                from: state.selection.head,
                to: state.selection.head,
                text: " ",
              });
              expect(prepared).toBeUndefined();
              expect(session.document).toStrictEqual(before);
              expect(state.selection.toJSON()).toStrictEqual(selection);
            }
          }
        }
      }),
      { numRuns: 4 },
    );
  });

  test("generated restart, continue and nesting commands preserve zero starts and exact history", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(
          fc.record({
            start: fc.integer({ min: 0, max: 30 }),
            depth: fc.integer({ min: 1, max: 3 }),
          }),
          { minLength: 2, maxLength: 5 },
        ),
        async (generated) => {
          const document = seed();
          for (const [index, paragraph] of document.package.document.content.entries()) {
            if (paragraph.type === "paragraph")
              assignParagraphPropertySource(
                paragraph,
                `<w:pPr><w:rsidRPr w:val="${index}"/></w:pPr>`,
              );
          }
          assignDocumentParagraphPropertySourceContract(document, "b".repeat(64));
          const session = createCanonicalSession(document).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          const history: {
            before: Document;
            after: Document;
            preSelection: ReturnType<typeof state.selection.toJSON>;
            postSelection: ReturnType<typeof state.selection.toJSON>;
          }[] = [];
          const apply = (commands: Parameters<typeof prepareCanonicalCommands>[2]) => {
            const before = session.document;
            const preSelection = state.selection.toJSON();
            state = accept(
              session,
              state,
              prepareCanonicalCommands(session, state, commands).unwrap(),
            );
            history.push({
              before,
              after: session.document,
              preSelection,
              postSelection: state.selection.toJSON(),
            });
            expect(state.selection.toJSON()).toEqual(preSelection);
            expect(texts(session)).toEqual(["ab😀cd", "EF"]);
          };
          state = select(state, 1, state.doc.content.size - 1);
          apply([{ type: "toggleList", kind: "decimal" }]);
          const firstNumbering = paragraphs(session).at(0)?.formatting?.numPr;
          if (firstNumbering?.kind !== "reference")
            return panic("List creation omitted first numbering.");
          const firstId = firstNumbering.numId;
          const exercised = new Set<string>();
          for (const { start, depth } of [{ start: 0, depth: 2 }, ...generated]) {
            const second =
              paragraphs(session).at(1) ?? panic("Numbering sequence lost its second paragraph.");
            const address =
              session.projection.paragraph(second.paraId ?? "") ??
              panic("Numbering sequence lost its address.");
            state = select(state, address.start + 1, address.start);
            const count = session.document.package.numbering?.nums.length ?? 0;
            apply([{ type: "restartNumbering", start }]);
            exercised.add("restartNumbering");
            const restarted = paragraphs(session).at(1)?.formatting?.numPr;
            if (restarted?.kind !== "reference") return panic("Restart removed numbering.");
            expect(restarted.numId).not.toBe(firstId);
            expect(session.document.package.numbering?.nums.length).toBe(count + 1);
            expect(
              session.document.package.numbering?.nums
                .find(({ numId }) => numId === restarted.numId)
                ?.levelOverrides?.find(({ ilvl }) => ilvl === 0)?.startOverride,
            ).toBe(start);
            expect(state.doc.child(1).attrs["listStartOverride"]).toBe(start);
            for (let level = 1; level <= depth; level += 1) {
              apply([{ type: "changeListLevel", direction: "increase" }]);
              exercised.add("indent");
              const numbered = paragraphs(session).at(1)?.formatting?.numPr;
              expect(numbered?.kind === "reference" ? numbered.ilvl : undefined).toBe(level);
            }
            for (let level = depth - 1; level >= 0; level -= 1) {
              apply([{ type: "changeListLevel", direction: "decrease" }]);
              exercised.add("outdent");
              const numbered = paragraphs(session).at(1)?.formatting?.numPr;
              expect(numbered?.kind === "reference" ? numbered.ilvl : undefined).toBe(level);
            }
            apply([{ type: "continueNumbering" }]);
            exercised.add("continueNumbering");
            const continued = paragraphs(session).at(1)?.formatting?.numPr;
            expect(continued?.kind === "reference" ? continued.numId : undefined).toBe(firstId);
          }
          expect([...exercised].sort()).toEqual([
            "continueNumbering",
            "indent",
            "outdent",
            "restartNumbering",
          ]);
          const saved = await createDocx(session.document);
          const reopened = await parseDocx(saved, { preloadFonts: false });
          expect(reopened.package.numbering?.nums).toEqual(
            session.document.package.numbering?.nums,
          );
          expect(
            reopened.package.numbering?.nums.some(({ levelOverrides }) =>
              levelOverrides?.some(({ startOverride }) => startOverride === 0),
            ),
          ).toBe(true);
          for (const entry of history.toReversed()) {
            state = accept(session, state, session.prepareUndo(state).unwrap());
            expect(session.document).toStrictEqual(entry.before);
            expect(state.selection.toJSON()).toEqual(entry.preSelection);
          }
          expect(session.canUndo).toBe(false);
          for (const entry of history) {
            state = accept(session, state, session.prepareRedo(state).unwrap());
            expect(session.document).toStrictEqual(entry.after);
            expect(state.selection.toJSON()).toEqual(entry.postSelection);
          }
          expect(session.canRedo).toBe(false);
        },
      ),
      { numRuns: 12 },
    );
  });

  test("generated cross-paragraph selections format and replace across run boundaries with exact undo", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom(0, 1, 2, 4, 5),
        fc.integer({ min: 1, max: 2 }),
        fc.boolean(),
        fc.constantFrom("Z", "é", "😀"),
        async (firstOffset, secondOffset, reverse, replacement) => {
          const document = seed();
          const first = document.package.document.content.at(0);
          const second = document.package.document.content.at(1);
          if (first?.type !== "paragraph" || second?.type !== "paragraph")
            return panic("Cross-paragraph fixture lost its paragraphs.");
          first.content = [
            { type: "run", formatting: { italic: true }, content: [{ type: "text", text: "ab" }] },
            {
              type: "run",
              formatting: { underline: { style: "single" } },
              content: [{ type: "text", text: "😀cd" }],
            },
          ];
          const session = createCanonicalSession(document).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          const firstAddress =
            session.projection.paragraph(first.paraId ?? "") ??
            panic("Missing first selection address.");
          const secondAddress =
            session.projection.paragraph(second.paraId ?? "") ??
            panic("Missing second selection address.");
          const from = firstAddress.start + firstOffset;
          const to = secondAddress.start + secondOffset;
          state = select(state, reverse ? to : from, reverse ? from : to);
          const original = session.document;
          const originalSelection = state.selection.toJSON();
          state = accept(
            session,
            state,
            prepareCanonicalCommands(session, state, [
              { type: "formatRun", from, to, patch: { bold: true } },
            ]).unwrap(),
          );
          expect(state.selection.toJSON()).toEqual(originalSelection);
          state.doc.nodesBetween(from, to, (node) => {
            if (node.isText) expect(node.marks.some(({ type }) => type.name === "bold")).toBe(true);
          });
          const formatted = session.document;
          state = accept(
            session,
            state,
            session.prepareReplace(state, { from, to, text: replacement }).unwrap(),
          );
          const replaced = session.document;
          const expected = "ab😀cd".slice(0, firstOffset) + replacement + "EF".slice(secondOffset);
          expect(texts(session)).toEqual([expected]);
          expect(paragraphs(session).at(0)?.paraId).toBe(second.paraId);
          expect(session.projection.selectionAt(state).unwrap().head.offset).toBe(
            firstOffset + replacement.length,
          );
          const postSelection = state.selection.toJSON();
          state = accept(session, state, session.prepareUndo(state).unwrap());
          expect(session.document).toEqual(formatted);
          expect(state.selection.toJSON()).toEqual(originalSelection);
          state = accept(session, state, session.prepareUndo(state).unwrap());
          expect(session.document).toEqual(original);
          expect(state.selection.toJSON()).toEqual(originalSelection);
          state = accept(session, state, session.prepareRedo(state).unwrap());
          expect(session.document).toEqual(formatted);
          state = accept(session, state, session.prepareRedo(state).unwrap());
          expect(session.document).toEqual(replaced);
          expect(state.selection.toJSON()).toEqual(postSelection);
        },
      ),
      { numRuns: 24 },
    );
  });

  test("split identities never reuse retired identities and undo restores each source lineage", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(fc.constantFrom(0, 1, 2, 4, 5, 6, 7, 8), { minLength: 1, maxLength: 8 }),
        async (offsets) => {
          const document = seed();
          const first = document.package.document.content.at(0);
          const second = document.package.document.content.at(1);
          if (first?.type !== "paragraph" || second?.type !== "paragraph")
            return panic("Missing lineage fixture.");
          first.paraId = "00000001";
          second.paraId = "00000002";
          const xml = "<w:pPr><w:keepNext/></w:pPr>";
          assignParagraphPropertySource(first, xml);
          assignDocumentParagraphPropertySourceContract(document, "a".repeat(64));
          const token = getParagraphPropertySourceToken(first);
          const session = createCanonicalSession(document).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          state = select(state, 7);
          state = accept(session, state, session.prepareJoin(state, "forward").unwrap());
          const joined = session.document;
          const retired = new Set(["00000001"]);
          for (const offset of offsets) {
            state = select(state, offset + 1);
            state = accept(session, state, session.prepareSplit(state).unwrap());
            const allocated = paragraphs(session).find(({ paraId }) => paraId !== "00000002");
            if (allocated?.paraId === undefined) return panic("Split omitted its new identity.");
            expect(retired.has(allocated.paraId)).toBe(false);
            retired.add(allocated.paraId);
            expect(getParagraphPropertySource(allocated)).toBeUndefined();
            expect(getParagraphPropertySourceToken(allocated)).toBeUndefined();
            const split = session.document;
            state = accept(session, state, session.prepareUndo(state).unwrap());
            expect(session.document).toStrictEqual(joined);
            state = accept(session, state, session.prepareRedo(state).unwrap());
            expect(session.document).toStrictEqual(split);
            state = accept(session, state, session.prepareUndo(state).unwrap());
          }
          state = accept(session, state, session.prepareUndo(state).unwrap());
          const restored =
            paragraphs(session).at(0) ?? panic("Undo did not restore retired paragraph.");
          expect(restored.paraId).toBe("00000001");
          expect(getParagraphPropertySource(restored)?.xml).toBe(xml);
          expect(getParagraphPropertySourceToken(restored)).toBe(token);
          expect(texts(session)).toEqual(["ab😀cd", "EF"]);
        },
      ),
      { numRuns: 24 },
    );
  });

  test("numbering definitions survive DOCX save and reopen after list commands", async () => {
    const session = createCanonicalSession(seed()).unwrap();
    let state = EditorState.create({ schema, doc: session.projection.doc });
    state = select(state, 1);
    state = accept(
      session,
      state,
      prepareCanonicalCommands(session, state, [{ type: "toggleList", kind: "decimal" }]).unwrap(),
    );
    const authored = paragraphs(session).at(0)?.formatting?.numPr;
    expect(authored?.kind).toBe("reference");
    const buffer = await createDocx(session.document);
    const zip = await JSZip.loadAsync(buffer);
    expect(await zip.file("word/numbering.xml")?.async("string")).toContain(
      'w:numFmt w:val="decimal"',
    );
    const reopened = await parseDocx(buffer, { preloadFonts: false });
    expect(reopened.package.numbering?.nums).toEqual(session.document.package.numbering?.nums);
    const paragraph = reopened.package.document.content.at(0);
    expect(paragraph?.type === "paragraph" ? paragraph.formatting?.numPr : undefined).toEqual(
      authored,
    );
  });
});

const tabStopArbitrary = fc.record({
  position: fc.integer({ min: 0, max: 5000 }),
  alignment: fc.constantFrom(...TAB_STOP_ALIGNMENT_VALUES),
  leader: fc.constantFrom(...TAB_LEADER_VALUES),
});
const paragraphCommandArbitrary = fc.record({
  position: fc.integer({ min: 0, max: 5000 }),
  alignment: fc.constantFrom(...TAB_STOP_ALIGNMENT_VALUES),
  leader: fc.constantFrom(...TAB_LEADER_VALUES),
  tabs: fc.uniqueArray(tabStopArbitrary, { maxLength: 4, selector: (tab) => tab.position }),
  reverse: fc.boolean(),
  selection: fc.constantFrom("caret", "first", "all"),
});
type ParagraphCommandInput = ReturnType<typeof paragraphCommandArbitrary.generate>["value"];
const paragraphCommandFactories = {
  toggleBidi: () => singletonManager.requireCommand("toggleBidi")(),
  setRtl: () => singletonManager.requireCommand("setRtl")(),
  setLtr: () => singletonManager.requireCommand("setLtr")(),
  setTabs: ({ tabs }: ParagraphCommandInput) => singletonManager.requireCommand("setTabs")(tabs),
  addTabStop: ({ position, alignment, leader }: ParagraphCommandInput) =>
    singletonManager.requireCommand("addTabStop")(position, alignment, leader),
  removeTabStop: ({ position }: ParagraphCommandInput) =>
    singletonManager.requireCommand("removeTabStop")(position),
};

test("generated direction and tab command histories preserve authored values and exact inverse", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.tuple(fc.constantFrom(undefined, false, true), fc.constantFrom(undefined, false, true)),
      fc.tuple(
        fc.option(
          fc.uniqueArray(tabStopArbitrary, { maxLength: 4, selector: (tab) => tab.position }),
          { nil: undefined },
        ),
        fc.option(
          fc.uniqueArray(tabStopArbitrary, { maxLength: 4, selector: (tab) => tab.position }),
          { nil: undefined },
        ),
      ),
      fc.boolean(),
      fc.array(paragraphCommandArbitrary, { minLength: 6, maxLength: 12 }),
      async (directions, tabSets, empty, inputs) => {
        const document = seed();
        for (const [index, paragraph] of document.package.document.content.entries()) {
          if (paragraph.type !== "paragraph")
            return panic("Paragraph command seed has a nonparagraph.");
          const bidi = directions.at(index);
          const tabs = tabSets.at(index);
          paragraph.formatting = {
            ...(bidi === undefined ? {} : { bidi }),
            ...(tabs === undefined ? {} : { tabs }),
          };
          if (empty) paragraph.content = [];
        }
        for (const mode of [
          { type: "editing" },
          { type: "suggesting", author: "Reviewer" },
        ] as const) {
          const session = createCanonicalSession(document).unwrap();
          session.setMode(mode);
          let state = EditorState.create({ schema, doc: session.projection.doc });
          const baseline = session.document;
          const history: {
            before: Document;
            after: Document;
            preSelection: ReturnType<typeof state.selection.toJSON>;
            postSelection: ReturnType<typeof state.selection.toJSON>;
          }[] = [];
          const exercised = new Set<string>();
          const commands = Object.entries(paragraphCommandFactories);
          for (const [index, input] of inputs.entries()) {
            const [name, factory] =
              commands.at(index % commands.length) ?? panic("Missing paragraph command factory.");
            exercised.add(name);
            const first =
              session.projection.paragraph("12345678") ?? panic("Missing first paragraph address.");
            const last =
              session.projection.paragraph("00000001") ?? panic("Missing last paragraph address.");
            const from = first.start;
            let to = from;
            if (input.selection === "first") to = first.start + first.text.length;
            if (input.selection === "all") to = last.start + last.text.length;
            state = select(state, input.reverse ? to : from, input.reverse ? from : to);
            const command = factory(input);
            const before = session.document;
            const version = session.version;
            const preSelection = state.selection.toJSON();
            const storedMarks = state.storedMarks;
            command(state);
            const intents =
              getCanonicalCommandIntents(command, state) ?? panic(`Missing descriptor for ${name}`);
            expect(getCanonicalCommandIntents(command, state)).toEqual(intents);
            expect(session.document).toBe(before);
            expect(session.version).toBe(version);
            expect(state.selection.toJSON()).toEqual(preSelection);
            expect(state.storedMarks).toBe(storedMarks);
            const legacyBefore = state;
            let legacy = legacyBefore;
            expect(
              command(legacyBefore, (transaction) => {
                legacy = legacyBefore.apply(transaction);
              }),
            ).toBe(true);
            // Compare the command's current attrs, not PM-to-model serialization:
            // imported paragraphs retain original tabs in that legacy serializer.
            const authored: {
              bidi: boolean | undefined;
              tabs: ReturnType<typeof expectParagraphAttrs>["tabs"] | undefined;
            }[] = [];
            legacy.doc.forEach((paragraph) => {
              const attrs = expectParagraphAttrs(paragraph);
              authored.push({
                bidi: directionToAuthoredBidi(attrs.direction),
                tabs: attrs.tabs ?? undefined,
              });
            });
            session.breakUndoGroup();
            const prepared = prepareCanonicalCommands(session, state, intents);
            if (prepared.isErr()) {
              expect(prepared.error.reason).toBe("noChange");
              expect(session.document).toBe(before);
              expect(session.version).toBe(version);
              expect(state.selection.toJSON()).toEqual(preSelection);
            } else {
              state = accept(session, state, prepared.value);
              history.push({
                before,
                after: session.document,
                preSelection,
                postSelection: state.selection.toJSON(),
              });
            }
            const ordered = (formatting: (typeof authored)[number]) => ({
              ...formatting,
              tabs: formatting.tabs?.toSorted((a, b) => a.position - b.position),
            });
            expect(
              paragraphs(session).map((paragraph) =>
                ordered({ bidi: paragraph.formatting?.bidi, tabs: paragraph.formatting?.tabs }),
              ),
            ).toEqual(authored.map(ordered));
            expect(state.selection.toJSON()).toEqual(legacy.selection.toJSON());
          }
          expect([...exercised].sort()).toEqual(Object.keys(paragraphCommandFactories).sort());
          const saved = await parseDocx(await createDocx(session.document));
          // OOXML omits empty tabs and the default leader; compare their semantic values.
          // History below separately checks exact authored arrays and optional fields.
          const emittedFormatting = (
            paragraph: Document["package"]["document"]["content"][number],
          ) => {
            if (paragraph.type !== "paragraph") return panic("Reopen lost a paragraph.");
            return {
              bidi: paragraph.formatting?.bidi,
              tabs: (paragraph.formatting?.tabs ?? [])
                .toSorted((a, b) => a.position - b.position)
                .map((tab) => Object.assign({}, tab, { leader: tab.leader ?? "none" })),
            };
          };
          expect(saved.package.document.content.map(emittedFormatting)).toEqual(
            paragraphs(session).map(emittedFormatting),
          );
          for (const entry of history.toReversed()) {
            state = accept(session, state, session.prepareUndo(state).unwrap());
            expect(session.document).toStrictEqual(entry.before);
            expect(state.selection.toJSON()).toEqual(entry.preSelection);
          }
          expect(session.document).toStrictEqual(baseline);
          for (const entry of history) {
            state = accept(session, state, session.prepareRedo(state).unwrap());
            expect(session.document).toStrictEqual(entry.after);
            expect(state.selection.toJSON()).toEqual(entry.postSelection);
          }
        }
      },
    ),
    { numRuns: 12 },
  );
});
