import { describe, expect, test, setDefaultTimeout, spyOn } from "bun:test";
import { EditorState, Plugin, TextSelection } from "prosemirror-state";

import { schema } from "../prosemirror/schema";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { normalizeForOps, DOCUMENT_OP_TYPES, OP_STORIES } from "@stll/docx-core/ops";
import * as documentOps from "@stll/docx-core/ops";
import * as conversion from "../prosemirror/conversion/toProseDoc";
import { panic } from "better-result";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { FIRST_ID, fixture, seedArbitrary } from "../../typecheck/ops/reviewGenerators.typecheck";

setDefaultTimeout(propertyTestTimeout(240_000));
import {
  assignDocumentParagraphPropertySourceContract,
  assignParagraphPropertySource,
  getDocumentParagraphPropertySourceContract,
  getParagraphPropertySource,
  getParagraphPropertySourceToken,
  paragraphPropertySourceBelongsToDocument,
} from "../docx/paragraphPropertySource";
import type { Document, Run, StyleDefinitions } from "../types/document";
import {
  CANONICAL_PROJECTION_META,
  createCanonicalSession,
  deletionRange,
  isCanonicalProjectionTransaction,
  publishCanonicalProjection,
  type CanonicalCommit,
  type CanonicalSession,
} from "./canonicalSession";

const seed = (text = "ab😀cd"): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          textId: "87654321",
          preservedAttributes: [{ name: "rsidR", value: "00112233" }],
          content: [
            {
              type: "run",
              formatting: { bold: true },
              preservedAttributes: [{ name: "rsidR", value: "00445566" }],
              content: [{ type: "text", text }],
            },
          ],
        },
        { type: "paragraph", paraId: "23456789", content: [] },
      ],
    },
  },
});

const stateFor = (session: CanonicalSession) =>
  EditorState.create({ schema, doc: session.projection.doc });

const accept = (state: EditorState, commit: CanonicalCommit) => {
  const nextState = state.apply(commit.transaction);
  expect(nextState.doc.eq(commit.projection.doc)).toBe(true);
  expect(commit.publish().isOk()).toBe(true);
  return nextState;
};

describe("canonical session", () => {
  test("canonical input sequences keep the projection, inverse history and refusal atomic", async () => {
    await assertProperty(
      fc.asyncProperty(
        seedArbitrary,
        fc.array(seedArbitrary, { minLength: 18, maxLength: 24 }),
        async (generated, inputs) => {
          const parsed = await fixture({ ...generated, container: "body", nesting: "plain" });
          const session = createCanonicalSession(
            normalizeForOps({
              ...parsed,
              package: {
                ...parsed.package,
                document: {
                  ...parsed.package.document,
                  sections: undefined,
                  content: parsed.package.document.content
                    .filter((block) => block.type === "paragraph")
                    .map((paragraph) =>
                      Object.assign({}, paragraph, {
                        content: [...generated.text].map(
                          (text, index) =>
                            ({
                              type: "run",
                              formatting: index % 2 === 0 ? { bold: true } : { italic: true },
                              content: [{ type: "text", text }],
                            }) satisfies Run,
                        ),
                      }),
                    ),
                },
              },
            }),
          ).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          const initial = session.document;
          const journal: {
            before: Document;
            after: Document;
            preSelection: ReturnType<typeof state.selection.toJSON>;
            postSelection: ReturnType<typeof state.selection.toJSON>;
          }[] = [];
          const kinds = [
            "insert",
            "delete",
            "replace",
            "reject",
            "op-refusal",
            "projection-refusal",
            "plugin-filter",
            "plugin-append",
            "plugin-throw",
          ] as const;
          const exercised = new Set<(typeof kinds)[number]>();
          for (const [index, input] of inputs.entries()) {
            const kind = kinds.at(index % kinds.length) ?? panic("Missing generated input kind");
            exercised.add(kind);
            const paragraph =
              session.projection.paragraph(FIRST_ID) ?? panic("Generated paragraph disappeared");
            const gaps = [paragraph.start];
            let offset = paragraph.start;
            for (const character of paragraph.text) {
              offset += character.length;
              gaps.push(offset);
            }
            const anchorIndex = input.offset % gaps.length;
            let headIndex = (input.offset + input.text.length) % gaps.length;
            if (headIndex === anchorIndex && gaps.length > 1)
              headIndex = (headIndex + 1) % gaps.length;
            const anchor = gaps.at(anchorIndex) ?? panic("Generated anchor disappeared");
            const head =
              kind === "insert"
                ? anchor
                : (gaps.at(headIndex) ?? panic("Generated head disappeared"));
            const from = Math.min(anchor, head);
            const to = Math.max(anchor, head);
            state = state.apply(
              state.tr.setSelection(
                TextSelection.create(
                  state.doc,
                  input.italic ? head : anchor,
                  input.italic ? anchor : head,
                ),
              ),
            );
            const before = session.document;
            const preSelection = state.selection.toJSON();
            const version = session.version;
            if (kind !== "insert" && kind !== "delete" && kind !== "replace") {
              const projection = session.projection;
              const history = { undo: session.canUndo, redo: session.canRedo };
              const validInput = { from, to, text: input.insertion };
              switch (kind) {
                case "reject": {
                  const invalidText = input.offset % 2 === 0 ? `${input.insertion}\n` : "\ud800";
                  expect(
                    session.prepareReplace(state, { from, to, text: invalidText }).isErr(),
                  ).toBe(true);
                  break;
                }
                case "op-refusal": {
                  const apply = documentOps.applyDocumentOps;
                  const injected = spyOn(documentOps, "applyDocumentOps").mockImplementation(
                    (document, ops) =>
                      apply(document, [
                        ...ops,
                        {
                          type: DOCUMENT_OP_TYPES.DELETE_RANGE,
                          from: {
                            story: OP_STORIES.MAIN,
                            blockId: FIRST_ID,
                            offset: Number.MAX_SAFE_INTEGER - 1,
                          },
                          to: {
                            story: OP_STORIES.MAIN,
                            blockId: FIRST_ID,
                            offset: Number.MAX_SAFE_INTEGER,
                          },
                        },
                      ]),
                  );
                  try {
                    const prepared = session.prepareReplace(state, validInput);
                    expect(injected).toHaveBeenCalledTimes(1);
                    expect(prepared.isErr()).toBe(true);
                  } finally {
                    injected.mockRestore();
                  }
                  break;
                }
                case "projection-refusal": {
                  const injected = spyOn(conversion, "toProseDoc").mockReturnValue(
                    schema.node("doc", null, [
                      schema.node("paragraph", { paraId: FIRST_ID }, schema.text("foreign")),
                    ]),
                  );
                  try {
                    const prepared = session.prepareReplace(state, validInput);
                    expect(injected).toHaveBeenCalledTimes(1);
                    expect(prepared.isErr()).toBe(true);
                  } finally {
                    injected.mockRestore();
                  }
                  break;
                }
                case "plugin-filter":
                case "plugin-append":
                case "plugin-throw": {
                  const plugin = (() => {
                    switch (kind) {
                      case "plugin-filter":
                        return new Plugin({ filterTransaction: () => false });
                      case "plugin-append":
                        return new Plugin({
                          appendTransaction: (_transactions, _oldState, nextState) =>
                            nextState.tr.insertText(input.insertion, from),
                        });
                      case "plugin-throw":
                        return new Plugin({
                          filterTransaction: () => {
                            throw new TypeError("Generated plugin staging refusal");
                          },
                        });
                      default: {
                        const unreachable: never = kind;
                        return panic(`Unknown generated plugin ${unreachable}`);
                      }
                    }
                  })();
                  const stagedState = state.reconfigure({ plugins: [plugin] });
                  const commit = session.prepareReplace(stagedState, validInput).unwrap();
                  expect(
                    publishCanonicalProjection({ state: stagedState, commit, session }).isErr(),
                  ).toBe(true);
                  expect(stagedState.doc).toBe(state.doc);
                  expect(stagedState.selection.toJSON()).toEqual(preSelection);
                  break;
                }
                default: {
                  const unreachable: never = kind;
                  panic(`Unknown generated refusal ${unreachable}`);
                }
              }
              expect(session.document).toBe(before);
              expect(session.projection).toBe(projection);
              expect(session.version).toBe(version);
              expect(state.selection.toJSON()).toEqual(preSelection);
              expect({ undo: session.canUndo, redo: session.canRedo }).toEqual(history);
              continue;
            }
            const text = kind === "delete" ? "" : input.insertion;
            const commit = session.prepareReplace(state, { from, to, text }).unwrap();
            expect(session.document).toBe(before);
            expect(session.version).toBe(version);
            expect(commit.transaction.getMeta("addToHistory")).toBe(false);
            state = publishCanonicalProjection({ state, commit, session }).unwrap().state;
            expect(session.version).toBe(version + 1);
            expect(state.doc.eq(toProseDoc(session.document))).toBe(true);
            expect(state.doc.eq(session.projection.doc)).toBe(true);
            expect(state.selection.from).toBe(from + text.length);
            expect(state.selection.to).toBe(from + text.length);
            journal.push({
              before,
              after: session.document,
              preSelection,
              postSelection: state.selection.toJSON(),
            });
          }
          expect([...exercised].sort()).toEqual([...kinds].sort());
          for (const entry of journal.toReversed()) {
            const version = session.version;
            const commit = session.prepareUndo(state).unwrap();
            state = publishCanonicalProjection({ state, commit, session }).unwrap().state;
            expect(session.document).toStrictEqual(entry.before);
            expect(state.selection.toJSON()).toEqual(entry.preSelection);
            expect(session.version).toBe(version + 1);
            expect(state.doc.eq(toProseDoc(session.document))).toBe(true);
          }
          expect(session.document).toStrictEqual(initial);
          expect(session.canUndo).toBe(false);
          for (const entry of journal) {
            const version = session.version;
            const commit = session.prepareRedo(state).unwrap();
            state = publishCanonicalProjection({ state, commit, session }).unwrap().state;
            expect(session.document).toStrictEqual(entry.after);
            expect(state.selection.toJSON()).toEqual(entry.postSelection);
            expect(session.version).toBe(version + 1);
            expect(state.doc.eq(toProseDoc(session.document))).toBe(true);
          }
          expect(session.canRedo).toBe(false);
        },
      ),
      { numRuns: 40 },
    );
  });

  test("host seed and style mutations cannot change the authority or its inverses", () => {
    const document = seed();
    const paragraph = document.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") panic("Seed paragraph unavailable");
    paragraph.formatting = { styleId: "HostStyle" };
    const authoredStyle = { italic: true };
    const hostStyles = {
      styles: [{ styleId: "HostStyle", type: "paragraph", rPr: authoredStyle }],
    } satisfies StyleDefinitions;
    const session = createCanonicalSession(document, hostStyles).unwrap();
    const initial = structuredClone(session.document);
    const initialProjection = session.projection.doc;
    const run = paragraph.content.at(0);
    if (run?.type !== "run") panic("Seed run unavailable");
    const leaf = run.content.at(0);
    if (leaf?.type !== "text") panic("Seed text unavailable");
    leaf.text = "host mutation";
    run.formatting = { bold: false };
    const hostStyle = hostStyles.styles.at(0);
    if (!hostStyle) panic("Host style unavailable");
    hostStyle.rPr.italic = false;
    expect(session.document).toStrictEqual(initial);
    expect(session.projection.doc).toBe(initialProjection);
    let state = stateFor(session);
    state = publishCanonicalProjection({
      state,
      session,
      commit: session.prepareReplace(state, { from: 1, to: 1, text: "X" }).unwrap(),
    }).unwrap().state;
    expect(
      state.doc.eq(
        toProseDoc(session.document, {
          styles: { styles: [{ styleId: "HostStyle", type: "paragraph", rPr: { italic: true } }] },
        }),
      ),
    ).toBe(true);
    state = publishCanonicalProjection({
      state,
      session,
      commit: session.prepareUndo(state).unwrap(),
    }).unwrap().state;
    expect(session.document).toStrictEqual(initial);
    expect(state.doc.eq(initialProjection)).toBe(true);
  });

  test("replacement stages immutable model/history and exact inverse restores IDs and selection", () => {
    const session = createCanonicalSession(seed()).unwrap();
    const original = session.document;
    let state = stateFor(session);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 6, 2)));
    const selected = state.selection.toJSON();
    const commit = session.prepareReplace(state, { from: 2, to: 6, text: "XY" }).unwrap();
    expect(session.document).toBe(original);
    expect(session.version).toBe(0);
    expect(session.canUndo).toBe(false);
    expect(commit.transaction.getMeta("addToHistory")).toBe(false);
    state = accept(state, commit);
    const edited = session.document;
    expect(state.doc.textContent).toBe("aXYd");
    expect(session.canUndo).toBe(true);
    expect(session.version).toBe(1);
    expect(state.selection.from).toBe(4);
    expect(commit.touched.modified).toEqual(["12345678"]);
    const unaffected = original.package.document.content.at(1);
    expect(edited.package.document.content.at(1)).toBe(unaffected);

    state = accept(state, session.prepareUndo(state).unwrap());
    expect(session.document).toEqual(original);
    expect(state.selection.toJSON()).toEqual(selected);
    expect(session.canRedo).toBe(true);
    expect(session.version).toBe(2);
    state = accept(state, session.prepareRedo(state).unwrap());
    expect(session.document).toEqual(edited);
    expect(state.doc.eq(session.projection.doc)).toBe(true);
    expect(state.selection.from).toBe(4);
    expect(session.version).toBe(3);
  });

  test("activation rejects absent/duplicate identities, secondary stories, wrappers and atoms", () => {
    const unsupported: Document[] = [
      { package: { document: { content: [{ type: "paragraph", content: [] }] } } },
      {
        package: {
          document: {
            content: [
              { type: "paragraph", paraId: "12345678", content: [] },
              { type: "paragraph", paraId: "12345678", content: [] },
            ],
          },
        },
      },
      {
        package: {
          ...seed().package,
          headers: new Map([["rId1", { type: "header", hdrFtrType: "default", content: [] }]]),
        },
      },
      {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                paraId: "12345678",
                content: [{ type: "run", content: [{ type: "tab" }] }],
              },
            ],
          },
        },
      },
      seed("bad\ud800"),
      seed("bad\u0000"),
    ];
    for (const document of unsupported) expect(createCanonicalSession(document).isErr()).toBe(true);
  });

  test("rejection leaves canonical model, projection, version and journal unchanged", () => {
    const session = createCanonicalSession(seed()).unwrap();
    const state = stateFor(session);
    const original = session.document;
    const projection = session.projection;
    const attempts = [
      { from: 4, to: 4, text: "X" }, // Inside the emoji's UTF-16 pair.
      { from: 2, to: 9, text: "X" }, // Cross-paragraph selection.
      { from: 2, to: 3, text: "\n" },
      { from: 2, to: 3, text: "\ud800" },
      { from: 2, to: 3, text: "\u0000" },
      { from: 3, to: 2, text: "X" },
    ];
    for (const input of attempts) expect(session.prepareReplace(state, input).isErr()).toBe(true);
    const stale = state.apply(state.tr.insertText("outside", 1));
    expect(session.prepareReplace(stale, { from: 1, to: 1, text: "X" }).isErr()).toBe(true);
    const bold = schema.marks.bold?.create();
    expect(bold).toBeDefined();
    if (bold === undefined) return;
    const formatted = state.apply(state.tr.setStoredMarks([bold]));
    expect(session.prepareReplace(formatted, { from: 1, to: 1, text: "X" }).isErr()).toBe(true);
    expect(session.document).toBe(original);
    expect(session.projection).toBe(projection);
    expect(session.version).toBe(0);
    expect(session.canUndo).toBe(false);
    expect(session.canRedo).toBe(false);
  });

  test("authorization uses private identity, expires after publish and rejects another session", () => {
    const session = createCanonicalSession(seed()).unwrap();
    const other = createCanonicalSession(seed()).unwrap();
    const state = stateFor(session);
    const commit = session.prepareReplace(state, { from: 1, to: 1, text: "X" }).unwrap();
    const spoofed = state.tr
      .insertText("outside", 1)
      .setMeta(CANONICAL_PROJECTION_META, { type: "canonical", origin: "input", version: 1 });
    expect(isCanonicalProjectionTransaction(spoofed, session)).toBe(false);
    expect(isCanonicalProjectionTransaction(commit.transaction, other)).toBe(false);
    expect(isCanonicalProjectionTransaction(commit.transaction, session)).toBe(true);
    accept(state, commit);
    expect(isCanonicalProjectionTransaction(commit.transaction, session)).toBe(false);
    expect(commit.publish().isErr()).toBe(true);
  });

  test("staged commits become stale and new input after undo invalidates redo", () => {
    const session = createCanonicalSession(seed()).unwrap();
    let state = stateFor(session);
    const first = session.prepareReplace(state, { from: 1, to: 1, text: "X" }).unwrap();
    const stale = session.prepareReplace(state, { from: 1, to: 1, text: "Y" }).unwrap();
    state = accept(state, first);
    expect(stale.publish().isErr()).toBe(true);
    state = accept(state, session.prepareUndo(state).unwrap());
    expect(session.canRedo).toBe(true);
    state = accept(state, session.prepareReplace(state, { from: 1, to: 1, text: "Z" }).unwrap());
    expect(session.canRedo).toBe(false);
    expect(session.prepareRedo(state).isErr()).toBe(true);
    expect(state.doc.textContent).toBe("Zab😀cd");
  });

  test("plain address map round-trips every code-point gap including empty paragraphs", () => {
    const session = createCanonicalSession(seed()).unwrap();
    for (const position of [1, 2, 3, 5, 6, 7, 9]) {
      const address = session.projection.addressAt(position).unwrap();
      expect(session.projection.positionAt(address).unwrap()).toBe(position);
    }
    expect(session.projection.addressAt(4).isErr()).toBe(true);
  });

  test("mutating an authorized projection transaction invalidates it atomically", () => {
    const session = createCanonicalSession(seed()).unwrap();
    const original = session.document;
    const state = stateFor(session);
    const commit = session.prepareReplace(state, { from: 1, to: 1, text: "X" }).unwrap();
    expect(isCanonicalProjectionTransaction(commit.transaction, session)).toBe(true);
    commit.transaction.insertText("outside", 1);
    expect(isCanonicalProjectionTransaction(commit.transaction, session)).toBe(false);
    expect(commit.publish().isErr()).toBe(true);
    expect(session.document).toBe(original);
    expect(session.version).toBe(0);
    expect(session.canUndo).toBe(false);
  });

  test("replacement inherits the first replaced authored run even when it is removed", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "12345678",
              content: [
                {
                  type: "run",
                  formatting: { bold: true },
                  content: [{ type: "text", text: "ab" }],
                },
                {
                  type: "run",
                  formatting: { italic: true },
                  content: [{ type: "text", text: "cd" }],
                },
              ],
            },
          ],
        },
      },
    };
    const session = createCanonicalSession(document).unwrap();
    let state = stateFor(session);
    state = accept(state, session.prepareReplace(state, { from: 1, to: 3, text: "X" }).unwrap());
    expect(state.doc.textContent).toBe("Xcd");
    expect(state.doc.resolve(1).nodeAfter?.marks.some((mark) => mark.type.name === "bold")).toBe(
      true,
    );
    state = accept(state, session.prepareUndo(state).unwrap());
    expect(session.document.package.document.content).toEqual(document.package.document.content);
  });

  test("normalization, text replacement, undo and redo preserve private paragraph-property provenance", () => {
    const document = seed();
    const paragraph = document.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") throw new TypeError("The fixture needs a paragraph.");
    const xml = "<w:pPr><w:keepNext/><w:contextualSpacing/></w:pPr>";
    assignParagraphPropertySource(paragraph, xml);
    assignDocumentParagraphPropertySourceContract(document, "a".repeat(64));
    paragraph.content.push({ type: "run", content: [] }); // Forces normalization to replace the paragraph.
    const contract = getDocumentParagraphPropertySourceContract(document);
    const token = getParagraphPropertySourceToken(paragraph);
    const session = createCanonicalSession(document).unwrap();
    let state = stateFor(session);
    const check = () => {
      const current = session.document.package.document.content.at(0);
      if (current?.type !== "paragraph") throw new TypeError("The session lost its paragraph.");
      expect(getParagraphPropertySource(current)?.xml).toBe(xml);
      expect(getParagraphPropertySourceToken(current)).toBe(token);
      expect(getDocumentParagraphPropertySourceContract(session.document)).toBe(contract);
      expect(paragraphPropertySourceBelongsToDocument(current, session.document)).toBe(true);
    };
    check();
    state = accept(state, session.prepareReplace(state, { from: 2, to: 3, text: "X" }).unwrap());
    check();
    state = accept(state, session.prepareUndo(state).unwrap());
    check();
    state = accept(state, session.prepareRedo(state).unwrap());
    check();
    expect(state.doc.eq(session.projection.doc)).toBe(true);
  });

  test("deletion consumes emoji and combining graphemes without joining paragraphs", () => {
    const session = createCanonicalSession(seed("A😀éB")).unwrap();
    let state = stateFor(session);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 4)));
    expect(deletionRange(state, "backward").unwrap()).toEqual({ from: 2, to: 4 });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2)));
    expect(deletionRange(state, "forward").unwrap()).toEqual({ from: 2, to: 4 });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 6)));
    expect(deletionRange(state, "backward").unwrap()).toEqual({ from: 4, to: 6 });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 5)));
    expect(deletionRange(state, "backward").isErr()).toBe(true);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1)));
    expect(deletionRange(state, "backward").isErr()).toBe(true);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 7)));
    expect(deletionRange(state, "forward").isErr()).toBe(true);
  });
});
