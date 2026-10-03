import { describe, expect, test, setDefaultTimeout, spyOn } from "bun:test";
import { EditorState, Plugin, TextSelection } from "prosemirror-state";

import {
  createParagraphChangeTrackerPlugin,
  hasUntrackedChanges,
  clearTrackedChanges,
  paragraphChangeTrackerKey,
} from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { schema } from "../prosemirror/schema";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import {
  normalizeForOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  paragraphVisibleText,
  editorParagraphGroups,
  physicalOffsetAtVisibleOffset,
} from "@stll/docx-core/ops";
import * as documentOps from "@stll/docx-core/ops";
import * as conversion from "../prosemirror/conversion/toProseDoc";
import { panic } from "better-result";
import fc from "fast-check";
import { canonicalReviewBlocks } from "../../../../test/reviewProjection";
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
import type { Document, Paragraph, Run, StyleDefinitions } from "../types/document";
import {
  CANONICAL_PROJECTION_META,
  CanonicalSessionError,
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
  test("keeps typing groups separate from structural edits and suggesting mode", () => {
    const session = createCanonicalSession(seed("AB")).unwrap();
    let state = stateFor(session);
    const initial = session.document;
    for (const [time, text] of [
      [1000, "X"],
      [1001, "Y"],
    ] as const) {
      state = accept(
        state,
        session
          .prepareReplace(state, {
            from: state.selection.head,
            to: state.selection.head,
            text,
            semantic: "typing",
            time,
          })
          .unwrap(),
      );
    }
    const typed = session.document;
    state = accept(state, session.prepareSplit(state).unwrap());
    state = accept(state, session.prepareUndo(state).unwrap());
    expect(session.document).toEqual(typed);
    session.setMode({ type: "suggesting", author: "Reviewer" });
    for (const [time, text] of [
      [1002, "Z"],
      [1003, "W"],
    ] as const) {
      session.setMode({ type: "suggesting", author: "Reviewer" });
      state = accept(
        state,
        session
          .prepareReplace(state, {
            from: state.selection.head,
            to: state.selection.head,
            text,
            semantic: "typing",
            time,
          })
          .unwrap(),
      );
    }
    state = accept(state, session.prepareUndo(state).unwrap());
    expect(session.document).toEqual(typed);
    state = accept(state, session.prepareUndo(state).unwrap());
    expect(session.document).toEqual(initial);
    expect(state.doc.eq(session.projection.doc)).toBe(true);
  });

  test("separator note types, rather than producer-specific ids, govern activation", () => {
    for (const kind of ["footnote", "endnote"] as const) {
      for (const id of [0, 1, 7]) {
        for (const noteType of ["separator", "continuationSeparator", "normal"] as const) {
          const document = seed("Body");
          const content: Paragraph[] = [
            {
              type: "paragraph",
              paraId: "34567890",
              content: [
                {
                  type: "run",
                  content: [
                    {
                      type: "preservedXml",
                      xml: `<w:${noteType === "continuationSeparator" ? "continuationSeparator" : "separator"} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
                      text: "",
                    },
                  ],
                },
              ],
            },
          ];
          if (kind === "footnote")
            document.package.footnotes = [{ type: "footnote", id, noteType, content }];
          else document.package.endnotes = [{ type: "endnote", id, noteType, content }];
          expect(createCanonicalSession(document).isOk()).toBe(noteType !== "normal");
        }
      }
    }
  });

  test("secondary typing projects only its changed story and never serves stale cached content", () => {
    const document = seed("Body");
    document.package.headers = new Map([
      [
        "rIdHeader1",
        {
          type: "header",
          hdrFtrType: "default",
          content: [
            {
              type: "paragraph",
              paraId: "34567890",
              content: [{ type: "run", content: [{ type: "text", text: "Header" }] }],
            },
          ],
        },
      ],
    ]);
    const session = createCanonicalSession(document).unwrap();
    const header = { kind: "header", rId: "rIdHeader1" } as const;
    const bodySpy = spyOn(conversion, "toProseDoc");
    const storySpy = spyOn(conversion, "headerFooterToProseDoc");
    try {
      const initialProjection = session.projectStory(header).unwrap();
      expect(session.projectStory(header).unwrap()).toBe(initialProjection);
      expect(storySpy).toHaveBeenCalledTimes(1);
      const initialState = EditorState.create({ schema, doc: initialProjection.doc });
      const body = session.projection;
      const edit = session
        .prepareReplace(initialState, { from: 1, to: 1, text: "X", story: header })
        .unwrap();
      const state = accept(initialState, edit);
      expect(edit.bodyProjection).toBe(body);
      expect(session.projection).toBe(body);
      expect(bodySpy).not.toHaveBeenCalled();
      expect(storySpy).toHaveBeenCalledTimes(2);
      expect(session.projectStory(header).unwrap()).toBe(edit.projection);
      expect(session.projectStory(header).unwrap().doc.textContent).toBe("XHeader");
      expect(
        session
          .prepareReplace(initialState, { from: 1, to: 1, text: "stale", story: header })
          .isErr(),
      ).toBe(true);
      expect(storySpy).toHaveBeenCalledTimes(2);
      const mainState = stateFor(session);
      accept(mainState, session.prepareReplace(mainState, { from: 1, to: 1, text: "Y" }).unwrap());
      expect(session.projectStory(header).unwrap()).toBe(edit.projection);
      expect(storySpy).toHaveBeenCalledTimes(2);
      accept(state, session.prepareUndo(state, header).unwrap());
      expect(session.projectStory(header).unwrap().doc.textContent).toBe("XHeader");
      const current = EditorState.create({
        schema,
        doc: session.projectStory(header).unwrap().doc,
      });
      accept(current, session.prepareUndo(current, header).unwrap());
      expect(session.projectStory(header).unwrap().doc.textContent).toBe("Header");
    } finally {
      bodySpy.mockRestore();
      storySpy.mockRestore();
    }
  });

  test("undo can remove the story containing the active editor", () => {
    const session = createCanonicalSession(seed("Body")).unwrap();
    const initial = structuredClone(session.document);
    const footer = { kind: "footer", rId: "rIdFooterFirst" } as const;
    const bodyState = stateFor(session);
    accept(
      bodyState,
      session
        .prepareOperations(bodyState, [
          {
            type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
            sectionIndex: 0,
            story: footer,
            referenceType: "first",
            content: [{ type: "paragraph", paraId: "34567890", content: [] }],
          },
        ])
        .unwrap(),
    );
    const footerState = EditorState.create({
      schema,
      doc: session.projectStory(footer).unwrap().doc,
    });
    const removed = session.prepareUndo(footerState, footer).unwrap();
    accept(footerState, removed);
    expect(session.document).toStrictEqual(initial);
    expect(session.projectStory(footer).isErr()).toBe(true);
    const restored = session.prepareRedo(stateFor(session)).unwrap();
    accept(stateFor(session), restored);
    expect(session.projectStory(footer).isOk()).toBe(true);
  });
  test("package-only canonical commits require full save until the saved tracker clears", () => {
    const session = createCanonicalSession(seed("Body")).unwrap();
    let state = EditorState.create({
      schema,
      doc: session.projection.doc,
      plugins: [createParagraphChangeTrackerPlugin()],
    });
    const text = session.prepareReplace(state, { from: 1, to: 1, text: "Text" }).unwrap();
    state = accept(state, text);
    expect(hasUntrackedChanges(state)).toBe(false);
    const properties = session
      .prepareOperations(state, [
        { type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS, sectionIndex: 0, patch: { titlePg: true } },
      ])
      .unwrap();
    expect(properties.transaction.docChanged).toBe(false);
    state = accept(state, properties);
    expect(hasUntrackedChanges(state)).toBe(true);
    state = state.apply(clearTrackedChanges(state));
    expect(hasUntrackedChanges(state)).toBe(false);
    state = accept(state, session.prepareUndo(state).unwrap());
    expect(hasUntrackedChanges(state)).toBe(true);
  });
  test.each([
    [OP_STORIES.MAIN, false],
    [{ kind: "header", rId: "rIdHeader1" }, true],
  ] as const)("block insertion in %s tracks the operation's story", (story, packageChange) => {
    const source = seed("Body");
    source.package.headers = new Map([
      [
        "rIdHeader1",
        {
          type: "header",
          hdrFtrType: "default",
          content: [{ type: "paragraph", paraId: "34567890", content: [] }],
        },
      ],
    ]);
    const session = createCanonicalSession(source).unwrap();
    const original = structuredClone(session.document);
    const state = stateFor(session);
    const commit = session
      .prepareOperations(state, [
        {
          type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
          story,
          at: { type: "after", blockId: story === OP_STORIES.MAIN ? "12345678" : "34567890" },
          blocks: [{ type: "paragraph", paraId: "45678901", content: [] }],
        },
      ])
      .unwrap();
    expect(commit.transaction.getMeta(paragraphChangeTrackerKey)).toBe(
      packageChange ? "package-change" : undefined,
    );
    const inserted = accept(state, commit);
    accept(inserted, session.prepareUndo(inserted).unwrap());
    expect(session.document).toStrictEqual(original);
  });
  test("story typing groups and pending composition share the body journal boundary", () => {
    const session = createCanonicalSession(seed("Body")).unwrap();
    let bodyState = stateFor(session);
    const header = { kind: "header", rId: "rIdHeader1" } as const;
    bodyState = accept(
      bodyState,
      session
        .prepareOperations(bodyState, [
          {
            type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
            sectionIndex: 0,
            story: header,
            referenceType: "default",
            content: [{ type: "paragraph", paraId: "34567890", content: [] }],
          },
        ])
        .unwrap(),
    );
    const created = session.document;
    let state = EditorState.create({ schema, doc: session.projectStory(header).unwrap().doc });

    for (const [time, text] of [
      [1000, "X"],
      [1001, "Y"],
    ] as const) {
      state = accept(
        state,
        session
          .prepareReplace(state, {
            from: state.selection.head,
            to: state.selection.head,
            story: header,

            text,
            semantic: "typing",
            time,
          })
          .unwrap(),
      );
    }
    const typed = session.document;
    session.beginComposition().unwrap();
    expect(() => session.document).toThrow(CanonicalSessionError);
    expect(
      session.prepareReplace(state, { from: 1, to: 1, text: "Z", story: header }).isErr(),
    ).toBe(true);
    expect(session.prepareUndo(state, header).isErr()).toBe(true);
    session.endComposition();
    state = accept(state, session.prepareUndo(state, header).unwrap());
    expect(session.document).toEqual(created);
    state = accept(state, session.prepareRedo(state, header).unwrap());
    expect(session.document).toEqual(typed);
    expect(state.doc.textContent).toBe("XY");
  });

  test("story creation and edits share exact canonical undo and redo", () => {
    const session = createCanonicalSession(seed("Body")).unwrap();
    const initial = structuredClone(session.document);
    let state = stateFor(session);
    const header = { kind: "header", rId: "rIdHeader1" } as const;
    state = accept(
      state,
      session
        .prepareOperations(state, [
          {
            type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
            sectionIndex: 0,
            story: header,
            referenceType: "default",
            content: [
              {
                type: "paragraph",
                paraId: "34567890",
                content: [{ type: "run", content: [{ type: "text", text: "Header" }] }],
              },
            ],
          },
        ])
        .unwrap(),
    );
    const headerProjection = session.projectStory(header).unwrap();
    let headerState = EditorState.create({ schema, doc: headerProjection.doc });
    headerState = accept(
      headerState,
      session
        .prepareReplace(headerState, { from: 1, to: 7, text: "Edited", story: header })
        .unwrap(),
    );
    expect(headerState.doc.textContent).toBe("Edited");
    expect(session.projection.doc.textContent).toBe("Body");
    state = stateFor(session);
    state = accept(
      state,
      session
        .prepareOperations(state, [
          {
            type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
            sectionIndex: 0,
            story: { kind: "footer", rId: "rIdFooter1" },
            referenceType: "first",
            content: [{ type: "paragraph", paraId: "45678901", content: [] }],
          },
        ])
        .unwrap(),
    );
    expect(session.document.package.document.finalSectionProperties?.titlePg).toBe(true);
    state = accept(
      state,
      session
        .prepareOperations(state, [
          {
            type: DOCUMENT_OP_TYPES.ADD_NOTE,
            at: { story: OP_STORIES.MAIN, blockId: "12345678", offset: 4 },
            note: {
              type: "footnote",
              id: 12,
              content: [
                {
                  type: "paragraph",
                  paraId: "56789012",
                  content: [{ type: "run", content: [{ type: "text", text: "Note" }] }],
                },
              ],
            },
          },
        ])
        .unwrap(),
    );
    const noteStory = { kind: "footnote", id: 12 } as const;
    const noteProjection = session.projectStory(noteStory).unwrap();
    let noteState = EditorState.create({ schema, doc: noteProjection.doc });
    const noteStart = noteProjection
      .positionAt({ story: noteStory, blockId: "56789012", offset: 0 })
      .unwrap();
    const noteEnd = noteProjection
      .positionAt({ story: noteStory, blockId: "56789012", offset: 4 })
      .unwrap();
    noteState = accept(
      noteState,
      session
        .prepareReplace(noteState, {
          from: noteStart,
          to: noteEnd,
          text: "Edited note",
          story: noteStory,
        })
        .unwrap(),
    );
    expect(noteState.doc.textContent).toBe("Edited note");
    const final = structuredClone(session.document);
    state = stateFor(session);
    for (let count = 0; count < 5; count++)
      state = accept(state, session.prepareUndo(state).unwrap());
    expect(session.document).toStrictEqual(initial);
    for (let count = 0; count < 5; count++)
      state = accept(state, session.prepareRedo(state).unwrap());
    expect(session.document).toStrictEqual(final);
    expect(session.projection.doc.eq(state.doc)).toBe(true);
  });
  test.each(["footnote", "endnote"] as const)(
    "%s marker gaps map exactly and insertion after a marker preserves it through history",
    (kind) => {
      const source = seed("Body");
      const content = [
        {
          type: "paragraph",
          paraId: "34567890",
          content: [
            { type: "run", content: [{ type: "noteMarker", kind }] },
            { type: "run", content: [{ type: "text", text: "😀Note" }] },
          ],
        },
      ] satisfies Paragraph[];
      if (kind === "footnote") source.package.footnotes = [{ type: kind, id: 12, content }];
      else source.package.endnotes = [{ type: kind, id: 12, content }];
      const session = createCanonicalSession(source).unwrap();
      const initial = structuredClone(session.document);
      const story = { kind, id: 12 };
      const projection = session.projectStory(story).unwrap();
      for (const position of [1, 2, 4, 5, 6, 7, 8]) {
        const address = projection.addressAt(position).unwrap();
        expect(projection.positionAt(address).unwrap()).toBe(position);
      }
      expect(projection.addressAt(3).isErr()).toBe(true);
      const afterMarker = projection.positionAt({ story, blockId: "34567890", offset: 0 }).unwrap();
      expect(afterMarker).toBe(2);
      const deletingState = (position: number) =>
        EditorState.create({
          schema,
          doc: projection.doc,
          selection: TextSelection.create(projection.doc, position),
        });
      expect(deletionRange(deletingState(2), "forward").unwrap()).toEqual({ from: 2, to: 4 });
      expect(deletionRange(deletingState(4), "backward").unwrap()).toEqual({ from: 2, to: 4 });
      expect(deletionRange(deletingState(2), "backward").isErr()).toBe(true);
      expect(deletionRange(deletingState(1), "forward").isErr()).toBe(true);
      let state = EditorState.create({
        schema,
        doc: projection.doc,
        selection: TextSelection.create(projection.doc, afterMarker),
      });
      state = accept(
        state,
        session.prepareReplace(state, { from: 2, to: 2, text: "x", story }).unwrap(),
      );
      expect(state.selection.from).toBe(3);
      expect(state.doc.firstChild?.firstChild?.type.name).toBe("noteMarker");
      expect(state.doc.textContent).toBe("x😀Note");
      const edited = structuredClone(session.document);
      state = accept(state, session.prepareUndo(state, story).unwrap());
      expect(session.document).toStrictEqual(initial);
      state = accept(state, session.prepareRedo(state, story).unwrap());
      expect(session.document).toStrictEqual(edited);
      expect(state.selection.from).toBe(3);
    },
  );

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

describe("canonical tracked input", () => {
  test("generated replacements lower tracked batches and resolution shares exact inverse history", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(
          fc.record({
            index: fc.nat(50),
            count: fc.nat(4),
            text: fc.constantFrom("", "x", "😀", "é"),
          }),
          { minLength: 8, maxLength: 18 },
        ),
        async (trace) => {
          const direct = createCanonicalSession(seed("A😀B")).unwrap();
          const tracked = createCanonicalSession(seed("A😀B")).unwrap();
          tracked.setMode({ type: "suggesting", author: "Author" });
          let directState = stateFor(direct);
          let trackedState = stateFor(tracked);
          const snapshots = [tracked.document];
          for (const input of trace) {
            const paragraph = tracked.document.package.document.content.at(0);
            if (paragraph?.type !== "paragraph") panic("Trace lost its paragraph");
            const gaps = [0];
            let offset = 0;
            for (const character of paragraphVisibleText(paragraph)) {
              offset += character.length;
              gaps.push(offset);
            }
            const startIndex = input.index % gaps.length;
            const from = gaps.at(startIndex) ?? panic("Trace lost its gap");
            const to =
              gaps.at(Math.min(startIndex + input.count, gaps.length - 1)) ??
              panic("Trace lost its endpoint");
            if (from === to && input.text === "") continue;
            directState = accept(
              directState,
              direct
                .prepareReplace(directState, { from: from + 1, to: to + 1, text: input.text })
                .unwrap(),
            );
            trackedState = accept(
              trackedState,
              tracked
                .prepareReplace(trackedState, {
                  from: physicalOffsetAtVisibleOffset(paragraph, from) + 1,
                  to: physicalOffsetAtVisibleOffset(paragraph, to) + 1,
                  text: input.text,
                })
                .unwrap(),
            );
            expect(trackedState.doc.eq(toProseDoc(tracked.document))).toBe(true);
            snapshots.push(tracked.document);
          }
          const suggested = tracked.document;
          const revisions = new Set<number>();
          const scan = (value: unknown): void => {
            if (Array.isArray(value)) {
              for (const item of value) scan(item);
              return;
            }
            if (typeof value !== "object" || value === null) return;
            if (
              "info" in value &&
              typeof value.info === "object" &&
              value.info !== null &&
              "id" in value.info &&
              typeof value.info.id === "number"
            )
              revisions.add(value.info.id);
            for (const child of Object.values(value)) scan(child);
          };
          scan(suggested.package.document.content);
          for (const resolution of ["accept", "reject"] as const) {
            const resolved = createCanonicalSession(suggested).unwrap();
            let state = stateFor(resolved);
            state = accept(
              state,
              resolved.prepareResolve(state, { revisionIds: [...revisions], resolution }).unwrap(),
            );
            const expected = resolution === "accept" ? direct.document : snapshots.at(0);
            if (expected === undefined) panic("The resolution oracle lost its baseline.");
            expect(canonicalReviewBlocks(resolved.document.package.document.content)).toEqual(
              canonicalReviewBlocks(expected.package.document.content),
            );
            state = accept(state, resolved.prepareUndo(state).unwrap());
            expect(resolved.document).toEqual(suggested);
            state = accept(state, resolved.prepareRedo(state).unwrap());
            expect(state.doc.eq(resolved.projection.doc)).toBe(true);
          }
          for (let index = snapshots.length - 2; index >= 0; index -= 1) {
            trackedState = accept(trackedState, tracked.prepareUndo(trackedState).unwrap());
            expect(tracked.document).toEqual(snapshots.at(index));
          }
          for (let index = 1; index < snapshots.length; index += 1) {
            trackedState = accept(trackedState, tracked.prepareRedo(trackedState).unwrap());
            expect(tracked.document).toEqual(snapshots.at(index));
          }
        },
      ),
      { numRuns: 30 },
    );
  });

  test("deleted text is skipped by subsequent character input and structural batches undo exactly", () => {
    const session = createCanonicalSession(seed("abcd")).unwrap();
    session.setMode({ type: "suggesting", author: "Author" });
    let state = stateFor(session);
    const baseline = session.document;
    state = accept(state, session.prepareReplace(state, { from: 2, to: 3, text: "" }).unwrap());
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2)));
    const next = deletionRange(state, "forward").unwrap();
    expect(next).toEqual({ from: 2, to: 4 });
    state = accept(state, session.prepareReplace(state, { ...next, text: "" }).unwrap());
    const beforeSplit = session.document;
    state = accept(state, session.prepareSplit(state).unwrap());
    const split = session.document;
    state = accept(state, session.prepareJoin(state, "backward").unwrap());
    state = accept(state, session.prepareUndo(state).unwrap());
    expect(session.document).toEqual(split);
    state = accept(state, session.prepareUndo(state).unwrap());
    expect(session.document).toEqual(beforeSplit);
    state = accept(state, session.prepareUndo(state).unwrap());
    state = accept(state, session.prepareUndo(state).unwrap());
    expect(session.document).toEqual(baseline);
  });

  test.each(["backward", "forward"] as const)(
    "a tracked join followed by %s deletion consumes the visible character across the removed mark",
    (direction) => {
      const document = seed("ab");
      const second = document.package.document.content.at(1);
      if (second?.type !== "paragraph") panic("The join fixture lost its second paragraph.");
      second.content = [{ type: "run", content: [{ type: "text", text: "cd" }] }];
      const session = createCanonicalSession(document).unwrap();
      const baseline = session.document;
      session.setMode({ type: "suggesting", author: "Author" });
      let state = stateFor(session);
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 5)));
      state = accept(state, session.prepareJoin(state, "backward").unwrap());
      const joined = session.document;
      const caret = direction === "backward" ? 5 : 3;
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, caret)));
      const range = deletionRange(state, direction).unwrap();
      expect(range).toEqual(direction === "backward" ? { from: 2, to: 5 } : { from: 3, to: 6 });
      state = accept(state, session.prepareReplace(state, { ...range, text: "" }).unwrap());
      expect(
        editorParagraphGroups(session.document, OP_STORIES.MAIN).map(({ text }) => text),
      ).toEqual([direction === "backward" ? "acd" : "abd"]);
      const suggested = session.document;
      const revisionIds = suggested.package.document.content.flatMap((block) => {
        if (block.type !== "paragraph") return [];
        return [
          ...(block.pPrMark === undefined ? [] : [block.pPrMark.info.id]),
          ...block.content.flatMap((item) => (item.type === "deletion" ? [item.info.id] : [])),
        ];
      });
      for (const resolution of ["accept", "reject"] as const) {
        const resolved = createCanonicalSession(suggested).unwrap();
        let resolvedState = stateFor(resolved);
        resolvedState = accept(
          resolvedState,
          resolved.prepareResolve(resolvedState, { revisionIds, resolution }).unwrap(),
        );
        if (resolution === "accept") {
          expect(
            editorParagraphGroups(resolved.document, OP_STORIES.MAIN).map(({ text }) => text),
          ).toEqual([direction === "backward" ? "acd" : "abd"]);
        } else {
          expect(resolved.document.package.document.content).toEqual(
            baseline.package.document.content,
          );
        }
        resolvedState = accept(resolvedState, resolved.prepareUndo(resolvedState).unwrap());
        expect(resolved.document).toEqual(suggested);
      }
      state = accept(state, session.prepareUndo(state).unwrap());
      expect(session.document).toEqual(joined);
      state = accept(state, session.prepareUndo(state).unwrap());
      expect(session.document).toEqual(baseline);
    },
  );
});
