import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { panic } from "better-result";
import { EditorState, TextSelection } from "prosemirror-state";
import { paragraphLogicalText, type TextPosition } from "@stll/docx-core/ops";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document } from "../types/document";
import { schema } from "../prosemirror/schema";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import {
  assignDocumentParagraphPropertySourceContract,
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
          const session = createCanonicalSession(seed()).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          let oracle = ["ab😀cd", "EF"];
          const history: {
            before: Document;
            after: Document;
            preSelection: ReturnType<typeof state.selection.toJSON>;
            postSelection: ReturnType<typeof state.selection.toJSON>;
          }[] = [];
          const exercised = new Set<(typeof kinds)[number]>();
          const inputs = [
            ...kinds.map((kind) => ({ kind, offset: 2, reverse: false, bold: true })),
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
            let prepare: CanonicalCommit;
            let expectedCaret: TextPosition | undefined;
            const before = session.document;
            let preSelection = state.selection.toJSON();
            switch (input.kind) {
              case "split": {
                prepare = session.prepareSplit(state).unwrap();
                expectedCaret = { story: "main", blockId: source.paraId ?? "", offset: 0 };
                const original = oracle.at(index) ?? panic("Text oracle lost a paragraph.");
                oracle.splice(index, 1, original.slice(0, offset), original.slice(offset));
                break;
              }
              case "join": {
                if (sources.length === 1) {
                  expect(
                    session.prepareJoin(select(state, address.start), "backward").isErr(),
                  ).toBe(true);
                  continue;
                }
                const firstIndex = input.offset % (sources.length - 1);
                const first = sources.at(firstIndex) ?? panic("Missing join paragraph.");
                const firstAddress =
                  session.projection.paragraph(first.paraId ?? "") ??
                  panic("Missing join address.");
                state = select(state, firstAddress.start + firstAddress.text.length);
                preSelection = state.selection.toJSON();
                prepare = session.prepareJoin(state, "forward").unwrap();
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
                ]).unwrap();
                break;
              }
              case "paragraph":
                prepare = prepareCanonicalCommands(session, state, [
                  {
                    type: "formatParagraph",
                    at: state.selection.head,
                    patch: { alignment: input.reverse ? "right" : "center" },
                  },
                ]).unwrap();
                break;
              case "atom": {
                const at = session.projection.addressAt(state.selection.head).unwrap();
                expectedCaret = { ...at, offset: at.offset + 1 };
                const atom = (() => {
                  if (input.offset % 3 === 0) return { type: "break", breakType: "page" } as const;
                  if (input.reverse) return { type: "break", breakType: "textWrapping" } as const;
                  return { type: "tab" } as const;
                })();
                prepare = session
                  .prepareIntent(state, {
                    type: "insertAtom",
                    from: at,
                    to: at,
                    atom,
                  })
                  .unwrap();
                const original = oracle.at(index) ?? panic("Text oracle lost an atom paragraph.");
                oracle[index] = original.slice(0, offset) + "\uFFFC" + original.slice(offset);
                break;
              }
              case "list":
                prepare = prepareCanonicalCommands(session, state, [
                  { type: "toggleList", kind: input.reverse ? "bullet" : "decimal" },
                ]).unwrap();
                break;
              case "input": {
                expectedCaret = { story: "main", blockId: source.paraId ?? "", offset: offset + 1 };
                prepare = session
                  .prepareReplace(state, {
                    from: state.selection.head,
                    to: state.selection.head,
                    text: "é",
                    semantic: "replacement",
                  })
                  .unwrap();
                const original = oracle.at(index) ?? panic("Text oracle lost an input paragraph.");
                oracle[index] = original.slice(0, offset) + "é" + original.slice(offset);
                break;
              }
              default: {
                const unreachable: never = input.kind;
                return panic(`Unknown sequence kind ${unreachable}`);
              }
            }
            state = accept(session, state, prepare);
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
              const level = session.document.package.numbering?.abstractNums.at(0)?.levels.at(0);
              expect(level?.start).toBe(generated.start);
              expect(level?.lvlText).toBe(`%1${generated.punctuation}`);
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

  test("retired identities reused by split retain the source lineage of each undo state", () => {
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
    // Joining retires 00000001; the next split allocates that same unused identity.
    state = select(state, 3);
    state = accept(session, state, session.prepareSplit(state).unwrap());
    const allocated = paragraphs(session).find(({ paraId }) => paraId === "00000001");
    expect(allocated).toBeDefined();
    expect(allocated && getParagraphPropertySource(allocated)).toBeUndefined();
    state = accept(session, state, session.prepareUndo(state).unwrap());
    state = accept(session, state, session.prepareUndo(state).unwrap());
    const restored = paragraphs(session).at(0) ?? panic("Undo did not restore retired paragraph.");
    expect(restored.paraId).toBe("00000001");
    expect(getParagraphPropertySource(restored)?.xml).toBe(xml);
    expect(getParagraphPropertySourceToken(restored)).toBe(token);
    expect(texts(session)).toEqual(["ab😀cd", "EF"]);
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
