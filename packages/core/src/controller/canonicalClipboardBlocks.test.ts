/**
 * Clipboard block content controls, custom XML elements and kept block
 * markup paste whole, undo back to the document as it was, and survive a
 * save; a copy that starts or ends inside a container pastes the content it
 * copied from it.
 */
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { type DeepReadonly, readModel } from "../types/documentModel";
import { prepareCanonicalPaste } from "./canonicalClipboard";

const paragraph = (text: string, paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const control = (): BlockContent => ({
  type: "blockSdt",
  properties: { sdtType: "richText", id: 41, tag: "clause" },
  content: [paragraph("Inside control", "1000A001")],
});
const customXml = (): BlockContent => ({
  type: "blockCustomXml",
  openingXml:
    '<w:customXml xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:element="clause">',
  closingXml: "</w:customXml>",
  content: [paragraph("Inside custom XML", "1000B001")],
});
const preserved = (): BlockContent => ({
  type: "preservedBlock",
  xml: '<w:permStart xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="5" w:edGrp="everyone"/>',
});

/** The text of each block, a container's paragraphs as one entry. */
const shapeOf = (blocks: readonly DeepReadonly<BlockContent>[]): unknown[] =>
  blocks.flatMap((block): unknown[] => {
    switch (block.type) {
      case "paragraph":
        return [
          block.content
            .flatMap((item) =>
              item.type === "run"
                ? item.content.flatMap((leaf) =>
                    leaf.type === "text" ? [leaf.text] : [],
                  )
                : [],
            )
            .join(""),
        ];
      case "blockSdt":
      case "blockCustomXml":
        return [{ [block.type]: shapeOf(block.content) }];
      case "preservedBlock":
        return [block.type];
      default:
        return [];
    }
  });

const pasteInto = (copied: BlockContent) => {
  const source = createCanonicalSession(
    documentOf([paragraph("Before", "10000001"), copied]),
  ).unwrap();
  const sourceDoc = source.projection.doc;
  // The whole copied block, from the first paragraph's end.
  const slice = sourceDoc.slice(
    sourceDoc.child(0).nodeSize,
    sourceDoc.content.size,
  );
  const session = createCanonicalSession(
    documentOf([
      paragraph("Target", "20000001"),
      paragraph("Tail", "20000002"),
    ]),
  ).unwrap();
  let state = EditorState.create({ doc: session.projection.doc });
  // The caret at the end of the first paragraph.
  state = state.apply(
    state.tr.setSelection(
      TextSelection.create(state.doc, state.doc.child(0).nodeSize - 1),
    ),
  );
  return {
    session,
    state,
    slice,
    sourceDocument: readModel(source.document),
  };
};

describe.each([
  { kind: "block content control", block: control },
  { kind: "custom XML element", block: customXml },
  { kind: "kept block markup", block: preserved },
])("pasting a copied $kind", ({ block }) => {
  test("places it whole, undoes, and survives a save", async () => {
    const { session, state, slice, sourceDocument } = pasteInto(block());
    const before = session.document;
    const prepared = prepareCanonicalPaste({
      session,
      state,
      slice,
      sourceDocument,
    });
    if (prepared.isErr()) throw prepared.error;
    let next = publishCanonicalProjection({
      session,
      state,
      commit: prepared.value,
    }).unwrap().state;
    const pasted = shapeOf(session.document.package.document.content);
    expect(pasted).toContainEqual(shapeOf([block()])[0]);
    expect(pasted.at(0)).toBe("Target");
    expect(pasted.at(-1)).toBe("Tail");
    const reopened = await parseDocx(
      await createDocx(readModel(session.document)),
      {
        preloadFonts: false,
      },
    );
    expect(shapeOf(reopened.package.document.content)).toEqual(pasted);
    next = publishCanonicalProjection({
      session,
      state: next,
      commit: session.prepareUndo(next).unwrap(),
    }).unwrap().state;
    expect(shapeOf(session.document.package.document.content)).toEqual(
      shapeOf(before.package.document.content),
    );
  });
});

test("a copy from inside a content control pastes the text it copied", () => {
  const source = createCanonicalSession(
    documentOf([paragraph("Before", "10000001"), control()]),
  ).unwrap();
  const doc = source.projection.doc;
  // From the start of the control's text to the document's end: open
  // through the control at the start.
  const start = doc.child(0).nodeSize + 2;
  const slice = doc.slice(start, doc.content.size);
  expect(slice.openStart).toBeGreaterThan(1);
  const session = createCanonicalSession(
    documentOf([paragraph("Target", "20000001")]),
  ).unwrap();
  let state = EditorState.create({ doc: session.projection.doc });
  state = state.apply(
    state.tr.setSelection(
      TextSelection.create(state.doc, state.doc.child(0).nodeSize - 1),
    ),
  );
  const prepared = prepareCanonicalPaste({
    session,
    state,
    slice,
    sourceDocument: readModel(source.document),
  });
  if (prepared.isErr()) throw prepared.error;
  publishCanonicalProjection({
    session,
    state,
    commit: prepared.value,
  }).unwrap();
  expect(shapeOf(session.document.package.document.content).at(0)).toBe(
    "TargetInside control",
  );
});

test("a dragged content control moves whole, and undoes", () => {
  const session = createCanonicalSession(
    documentOf([
      paragraph("Before", "10000001"),
      control(),
      paragraph("After", "10000003"),
      paragraph("Tail", "10000004"),
    ]),
  ).unwrap();
  const doc = session.projection.doc;
  const before = shapeOf(session.document.package.document.content);
  // From the end of "Before" to the start of "After": the control alone.
  const from = doc.child(0).nodeSize - 1;
  const to = doc.child(0).nodeSize + doc.child(1).nodeSize + 1;
  const slice = doc.slice(from, to);
  // Inside "Tail", after "Ta".
  const tailStart =
    doc.child(0).nodeSize + doc.child(1).nodeSize + doc.child(2).nodeSize + 1;
  const state = EditorState.create({ doc });
  const prepared = prepareCanonicalPaste({
    session,
    state,
    slice,
    moveSource: { from, to },
    moveTarget: tailStart + 2,
  });
  if (prepared.isErr()) throw prepared.error;
  const next = publishCanonicalProjection({
    session,
    state,
    commit: prepared.value,
  }).unwrap().state;
  expect(shapeOf(session.document.package.document.content)).toEqual([
    "BeforeAfter",
    "Ta",
    { blockSdt: ["Inside control"] },
    "il",
  ]);
  publishCanonicalProjection({
    session,
    state: next,
    commit: session.prepareUndo(next).unwrap(),
  }).unwrap();
  expect(shapeOf(session.document.package.document.content)).toEqual(before);
});
