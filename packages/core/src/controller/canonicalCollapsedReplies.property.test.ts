import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";
import { Window } from "happy-dom";
import { DOMParser, DOMSerializer } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import {
  applyDocumentOp,
  applyDocumentOps,
  OP_STORIES,
} from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { DOCUMENT_SHAPES } from "../__tests__/documentShapes";
import { commentAnchorIdsForNode } from "../ai-edits/comment-lifecycle";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import {
  commentIdsNamedIn,
  removeCommentAnchors,
} from "../prosemirror/commentRecords";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { CommentReferenceExtension } from "../prosemirror/extensions/nodes/CommentReferenceExtension";
import { expectRangeAnchorAttrs } from "../prosemirror/rangeAnchorAttrs";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph, ParagraphContent } from "../types/document";

const run = (text: string) =>
  ({
    type: "run",
    content: [{ type: "text", text }],
  }) satisfies ParagraphContent;

const compareMarkers = (
  a: Extract<
    ParagraphContent,
    { type: "commentRangeStart" | "commentRangeEnd" }
  >,
  b: Extract<
    ParagraphContent,
    { type: "commentRangeStart" | "commentRangeEnd" }
  >,
) => {
  if (a.type === b.type) return a.id - b.id;
  return a.type === "commentRangeStart" ? -1 : 1;
};

const documentFor = (markers: ParagraphContent[]) =>
  ({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "10000001",
            content: [run("L"), ...markers, run("R")],
          },
        ],
        comments: markers.flatMap((marker) =>
          marker.type === "commentRangeStart"
            ? [
                {
                  id: marker.id,
                  author: "Reviewer",
                  anchorKind: "point",
                  content: [
                    {
                      type: "paragraph",
                      paraId: (0x20000000 + marker.id).toString(16),
                      content: [run("Comment")],
                    },
                  ],
                },
              ]
            : [],
        ),
      },
    },
  }) satisfies Document;

const bodyParagraph = (document: Document): Paragraph => {
  const paragraph = document.package.document.content.at(0);
  return paragraph?.type === "paragraph"
    ? paragraph
    : panic("Missing source paragraph.");
};

test(
  "generated collapsed comment clusters preserve authored marker order and indivisible source gaps",
  () => {
    assertProperty(
      fc.property(
        fc.array(fc.tuple(fc.nat(30), fc.nat(30)), {
          minLength: 2,
          maxLength: 8,
        }),
        (priorities) => {
          const markers = priorities
            .flatMap(([a, b], id) => [
              {
                priority: Math.min(a, b),
                marker: { type: "commentRangeStart", id } as const,
              },
              {
                priority: Math.max(a, b),
                marker: { type: "commentRangeEnd", id } as const,
              },
            ])
            .sort(
              (a, b) =>
                a.priority - b.priority || compareMarkers(a.marker, b.marker),
            )
            .map(({ marker }) => marker);
          const source = documentFor([
            ...markers,
            ...priorities.map(
              (_, id) =>
                ({ type: "commentReference", id }) satisfies ParagraphContent,
            ),
          ]);
          const before = structuredClone(source);
          const session = createCanonicalSession(source).unwrap();
          const native = session.projection.doc;
          expect(
            bodyParagraph(
              fromProseDoc(native, source, {
                stylesheetSource: { type: "package" },
              }),
            ).content,
          ).toEqual(bodyParagraph(source).content);
          expect(structuredClone(source)).toEqual(before);
          const atoms: ParagraphContent[] = [];
          let ordinal = 0;
          native.descendants((node, position) => {
            if (node.type.name !== "rangeAnchor") return true;
            expect(node.isAtom).toBe(true);
            expect(node.nodeSize).toBe(1);
            const { markers: carried } = expectRangeAnchorAttrs(node);
            atoms.push(...carried);
            expect(
              session.projection.addressAt(position).unwrap(),
            ).toMatchObject({ offset: 1, zeroWidthBefore: ordinal });
            for (let inner = 1; inner < carried.length; inner += 1)
              expect(
                session.projection
                  .positionAt({
                    story: OP_STORIES.MAIN,
                    blockId: "10000001",
                    offset: 1,
                    zeroWidthBefore: ordinal + inner,
                  })
                  .isErr(),
              ).toBe(true);
            ordinal += carried.length;
            expect(
              session.projection.addressAt(position + 1).unwrap(),
            ).toMatchObject({ offset: 1, zeroWidthBefore: ordinal });
            return false;
          });
          expect(atoms).toEqual(markers);
          expect([...commentIdsNamedIn(native)].sort((a, b) => a - b)).toEqual(
            priorities.map((_, id) => id),
          );
          const copied = schema.nodeFromJSON(native.toJSON());
          expect(
            bodyParagraph(
              fromProseDoc(copied, source, {
                stylesheetSource: { type: "package" },
              }),
            ).content,
          ).toEqual(bodyParagraph(source).content);
        },
      ),
      {
        id: "generated collapsed comment clusters preserve authored marker order and indivisible source gaps",
        seed: 110592881,
        numRuns: 60,
      },
    );
  },
  propertyTestTimeout(30_000),
);

test(
  "actual replies to an imported collapsed comment preserve source, history and saved anchors",
  async () => {
    const shape =
      DOCUMENT_SHAPES.find(({ id }) => id === "invisible-markers") ??
      panic("Missing marker shape.");
    const source = await parseDocx(
      (await prepareCanonicalDocxInput(await shape.build())).unwrap(),
      { preloadFonts: false, detectVariables: false },
    );
    const before = structuredClone(source);
    const session = createCanonicalSession(source).unwrap();
    let state = EditorState.create({ doc: session.projection.doc });
    const reply = {
      type: "createComment",
      comment: {
        id: 103,
        author: "Reviewer",
        content: [
          { type: "paragraph", paraId: "70000067", content: [run("Reply")] },
        ],
      },
      anchor: { kind: "reply", parentId: 3 },
    } as const satisfies Parameters<typeof applyDocumentOp>[1];
    const applied = applyDocumentOp(source, reply).unwrap();
    expect(structuredClone(source)).toEqual(before);
    expect(
      applyDocumentOps(applied.document, applied.inverse).unwrap().document,
    ).toEqual(source);
    expect(
      applyDocumentOps(
        applied.document,
        JSON.parse(JSON.stringify(applied.inverse)),
      ).unwrap().document,
    ).toEqual(source);
    state = publishCanonicalProjection({
      state,
      session,
      commit: session.prepareOperations(state, [reply]).unwrap(),
    }).unwrap().state;
    expect(session.document).toEqual(applied.document);
    const host = session.document.package.document.content.at(2);
    if (host?.type !== "paragraph") panic("Missing collapsed-comment host.");
    expect(
      host.content.filter(
        (item) =>
          item.type === "commentRangeStart" || item.type === "commentRangeEnd",
      ),
    ).toEqual([
      { type: "commentRangeStart", id: 3 },
      { type: "commentRangeStart", id: 103 },
      { type: "commentRangeEnd", id: 3 },
      { type: "commentRangeEnd", id: 103 },
    ]);
    expect(
      fromProseDoc(state.doc, session.document, {
        stylesheetSource: { type: "package" },
      }).package.document.content,
    ).toEqual(session.document.package.document.content);
    const saved = await serializeCanonicalSave({
      snapshot: session.captureSaveSnapshot(),
      featureFlags: { selectiveSave: true },
    });
    const reopened = await parseDocx(saved.buffer, {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(reopened.package.document.content).toEqual(
      session.document.package.document.content,
    );
    expect(
      reopened.package.document.comments?.map(({ id, parentId }) => ({
        id,
        parentId,
      })),
    ).toEqual(
      session.document.package.document.comments?.map(({ id, parentId }) => ({
        id,
        parentId,
      })),
    );
    expect(createCanonicalSession(reopened).isOk()).toBe(true);
    state = publishCanonicalProjection({
      state,
      session,
      commit: session.prepareUndo(state).unwrap(),
    }).unwrap().state;
    expect(session.document).toEqual(source);
    publishCanonicalProjection({
      state,
      session,
      commit: session.prepareRedo(state).unwrap(),
    }).unwrap();
    expect(session.document).toEqual(applied.document);
  },
  propertyTestTimeout(30_000),
);

const interleaved = [
  { type: "commentRangeStart", id: 3 },
  { type: "commentRangeStart", id: 103 },
  { type: "commentRangeEnd", id: 3 },
  { type: "commentRangeEnd", id: 103 },
] as const;
const interleavedSource = () =>
  documentFor([
    ...interleaved,
    { type: "commentReference", id: 3 },
    { type: "commentReference", id: 103 },
  ]);

test("nested point clusters keep their exact source-container ownership", () => {
  for (const wrapper of [
    { type: "inlineWrapper", kind: "bidi", control: "embedding" },
    { type: "inlineSdt", properties: { sdtType: "richText", id: 42 } },
  ] as const) {
    const source = documentFor([
      { ...wrapper, content: [...interleaved] },
      { type: "commentReference", id: 3 },
      { type: "commentReference", id: 103 },
    ]);
    source.package.document.comments =
      interleavedSource().package.document.comments;
    const before = structuredClone(source);
    const session = createCanonicalSession(source).unwrap();
    expect(
      bodyParagraph(
        fromProseDoc(session.projection.doc, source, {
          stylesheetSource: { type: "package" },
        }),
      ).content,
    ).toEqual(bodyParagraph(source).content);
    expect(structuredClone(source)).toEqual(before);
  }
});

test("removing one comment from a collapsed cluster preserves every sibling occurrence", () => {
  const source = interleavedSource();
  const native = toProseDoc(source);
  const state = EditorState.create({ doc: native });
  const tr = state.tr;
  removeCommentAnchors(tr, new Set([103]));
  expect([...commentIdsNamedIn(tr.doc)]).toEqual([3]);
  expect(
    bodyParagraph(
      fromProseDoc(tr.doc, source, { stylesheetSource: { type: "package" } }),
    ).content,
  ).toEqual([
    run("L"),
    { type: "commentRangeStart", id: 3 },
    { type: "commentRangeEnd", id: 3 },
    { type: "commentReference", id: 3 },
    run("R"),
  ]);
  expect(tr.doc.content.size).toBe(native.content.size - 1);
  native.descendants((node) => {
    if (node.type.name === "rangeAnchor")
      expect(commentAnchorIdsForNode(node)).toEqual([3, 103]);
    return true;
  });
});

test("collapsed reply clusters remain atomic through DOM copy, native deletion and inverted transactions", () => {
  const source = interleavedSource();
  const doc = toProseDoc(source);
  const atom = doc.child(0).child(1);
  expect(atom.type.name).toBe("rangeAnchor");
  const window = new Window();
  // SAFETY: happy-dom supplies the browser DOM implementation ProseMirror consumes.
  const document = window.document as unknown as globalThis.Document;
  const host = document.createElement("div");
  host.append(
    DOMSerializer.fromSchema(schema).serializeNode(atom, { document }),
  );
  const copied = DOMParser.fromSchema(schema).parse(host);
  expect(expectRangeAnchorAttrs(copied.child(0).child(0)).markers).toEqual(
    interleaved,
  );
  const plugin =
    CommentReferenceExtension().onSchemaReady({ schema }).plugins?.at(0) ??
    panic("Missing reference integrity plugin.");
  const state = EditorState.create({ doc, plugins: [plugin] });
  const at = 2;
  const typed = state.applyTransaction(state.tr.insertText("typed", at)).state;
  const untouched = fromProseDoc(typed.doc, source, {
    stylesheetSource: { type: "package" },
  });
  expect(
    bodyParagraph(untouched).content.filter(
      (item) =>
        item.type === "commentRangeStart" || item.type === "commentRangeEnd",
    ),
  ).toEqual(interleaved);
  expect(bodyParagraph(untouched).content.at(0)).toEqual(run("Ltyped"));
  const deleted = state.applyTransaction(
    state.tr.delete(at, at + atom.nodeSize),
  );
  expect([...commentIdsNamedIn(deleted.state.doc)]).toEqual([]);
  expect(deleted.state.doc.textContent).toBe("LR");
  let restored = deleted.state;
  for (const transaction of deleted.transactions.toReversed()) {
    const inverse = restored.tr;
    for (const [index, step] of [...transaction.steps.entries()].toReversed())
      inverse.step(
        step.invert(
          transaction.docs[index] ?? panic("Missing inverse source."),
        ),
      );
    restored = restored.apply(inverse);
  }
  expect(restored.doc.eq(doc)).toBe(true);
  expect(
    bodyParagraph(
      fromProseDoc(restored.doc, source, {
        stylesheetSource: { type: "package" },
      }),
    ).content,
  ).toEqual(bodyParagraph(source).content);
});

test(
  "generated actual reply histories preserve collapsed source occurrences through parsing and exact reversal",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(fc.nat(8), { minLength: 1, maxLength: 5 }),
        async (parents) => {
          const source = await parseDocx(
            await createDocx(
              documentFor([
                { type: "commentRangeStart", id: 3 },
                { type: "commentRangeEnd", id: 3 },
                { type: "commentReference", id: 3 },
              ]),
            ),
            { preloadFonts: false, detectVariables: false },
          );
          const session = createCanonicalSession(source).unwrap();
          const before = structuredClone(source);
          let state = EditorState.create({ doc: session.projection.doc });
          const ids = [3];
          for (const [index, choice] of parents.entries()) {
            const id = 103 + index;
            const parentId =
              ids.at(choice % ids.length) ?? panic("Missing reply parent.");
            const op = {
              type: "createComment",
              anchor: { kind: "reply", parentId },
              comment: {
                id,
                author: "Reviewer",
                content: [
                  {
                    type: "paragraph",
                    paraId: (0x70000000 + id).toString(16),
                    content: [run(`Reply ${index}`)],
                  },
                ],
              },
            } as const satisfies Parameters<typeof applyDocumentOp>[1];
            const applied = applyDocumentOp(session.document, op).unwrap();
            expect(
              applyDocumentOps(
                applied.document,
                JSON.parse(JSON.stringify(applied.inverse)),
              ).unwrap().document,
            ).toEqual(session.document);
            state = publishCanonicalProjection({
              state,
              session,
              commit: session.prepareOperations(state, [op]).unwrap(),
            }).unwrap().state;
            expect(session.document).toEqual(applied.document);
            expect(
              bodyParagraph(
                fromProseDoc(state.doc, session.document, {
                  stylesheetSource: { type: "package" },
                }),
              ).content,
            ).toEqual(bodyParagraph(session.document).content);
            expect(
              [...commentIdsNamedIn(state.doc)].sort((a, b) => a - b),
            ).toEqual([...ids, id]);
            ids.push(id);
          }
          const final = session.document;
          const saved = await serializeCanonicalSave({
            snapshot: session.captureSaveSnapshot(),
            featureFlags: { selectiveSave: true },
          });
          const reopened = await parseDocx(saved.buffer, {
            preloadFonts: false,
            detectVariables: false,
          });
          expect(bodyParagraph(reopened).content).toEqual(
            bodyParagraph(final).content,
          );
          expect(
            reopened.package.document.comments?.map(({ id, parentId }) => ({
              id,
              parentId,
            })),
          ).toEqual(
            final.package.document.comments?.map(({ id, parentId }) => ({
              id,
              parentId,
            })),
          );
          expect(createCanonicalSession(reopened).isOk()).toBe(true);
          for (const _ of parents)
            state = publishCanonicalProjection({
              state,
              session,
              commit: session.prepareUndo(state).unwrap(),
            }).unwrap().state;
          expect(structuredClone(session.document)).toEqual(before);
          for (const _ of parents)
            state = publishCanonicalProjection({
              state,
              session,
              commit: session.prepareRedo(state).unwrap(),
            }).unwrap().state;
          expect(session.document).toEqual(final);
          expect(structuredClone(source)).toEqual(before);
        },
      ),
      {
        id: "generated actual reply histories preserve collapsed source occurrences through parsing and exact reversal",
        seed: 115879302,
        numRuns: 20,
      },
    );
  },
  propertyTestTimeout(30_000),
);
