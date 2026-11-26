/**
 * Clipboard block content controls, custom XML elements and kept block
 * markup paste whole, undo back to the document as it was, and survive a
 * save; a copy that starts or ends inside a container pastes the content it
 * copied from it.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import type { Node as ProseNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

import { relationshipIdOf } from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

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

/** A session over `content` and a state over its projection. */
const sessionOver = (content: BlockContent[], document?: Document) => {
  const session = createCanonicalSession(
    document ?? documentOf(content),
  ).unwrap();
  return {
    session,
    state: EditorState.create({ doc: session.projection.doc }),
  };
};

/** A target session with the caret at the end of its first paragraph. */
const caretTarget = () => {
  const session = createCanonicalSession(
    documentOf([
      paragraph("Target", "20000001"),
      paragraph("Tail", "20000002"),
    ]),
  ).unwrap();
  let state = EditorState.create({ doc: session.projection.doc });
  state = state.apply(
    state.tr.setSelection(
      TextSelection.create(state.doc, state.doc.child(0).nodeSize - 1),
    ),
  );
  return { session, state };
};

/** Where the paragraph with `text` starts its text. */
const textStart = (doc: ProseNode, text: string): number => {
  let found: number | undefined;
  doc.descendants((node, position) => {
    if (found !== undefined) return false;
    if (node.type.name === "paragraph" && node.textContent === text)
      found = position + 1;
    return true;
  });
  return found ?? panic(`Missing paragraph ${text}.`);
};

test("a copy from between a control's paragraphs pastes them as paragraphs", () => {
  const { session: source } = sessionOver([
    paragraph("Before", "10000001"),
    {
      type: "blockSdt",
      properties: { sdtType: "richText", id: 41 },
      content: [
        paragraph("Inside one", "1000A001"),
        paragraph("Inside two", "1000A002"),
      ],
    },
    paragraph("After", "10000003"),
  ]);
  const doc = source.projection.doc;
  // From between the control's paragraphs into "After": open through the
  // control alone at the start.
  const slice = doc.slice(
    textStart(doc, "Inside two") - 1,
    textStart(doc, "After") + 2,
  );
  expect(slice.openStart).toBe(1);
  expect(slice.content.firstChild?.type.name).toBe("blockSdt");
  const { session, state } = caretTarget();
  const prepared = prepareCanonicalPaste({
    session,
    state,
    slice,
    sourceDocument: readModel(source.document),
  });
  if (prepared.isErr()) throw prepared.error;
  publishCanonicalProjection({ session, state, commit: prepared.value });
  expect(shapeOf(session.document.package.document.content)).toEqual([
    "Target",
    "Inside two",
    "Af",
    "Tail",
  ]);
});

test("a copy from a cell of a table inside a control pastes a table of the cells it holds", () => {
  const cell = (text: string, paraId: string) => ({
    type: "tableCell" as const,
    content: [paragraph(text, paraId)],
  });
  const { session: source } = sessionOver([
    paragraph("Before", "10000001"),
    {
      type: "blockSdt",
      properties: { sdtType: "richText", id: 41 },
      content: [
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [cell("A1", "1000C001"), cell("B1", "1000C002")],
            },
            {
              type: "tableRow",
              cells: [cell("A2", "1000C003"), cell("B2", "1000C004")],
            },
          ],
        },
      ],
    },
    paragraph("After", "10000003"),
  ]);
  const doc = source.projection.doc;
  const slice = doc.slice(
    textStart(doc, "A2") + 1,
    textStart(doc, "After") + 2,
  );
  expect(slice.content.firstChild?.type.name).toBe("blockSdt");
  const { session, state } = caretTarget();
  const before = session.document;
  const prepared = prepareCanonicalPaste({
    session,
    state,
    slice,
    sourceDocument: readModel(source.document),
  });
  if (prepared.isErr()) throw prepared.error;
  const next = publishCanonicalProjection({
    session,
    state,
    commit: prepared.value,
  }).unwrap().state;
  const content = session.document.package.document.content;
  const table = content.find((block) => block.type === "table");
  if (table?.type !== "table") throw new TypeError("Missing pasted table.");
  expect(
    table.rows.map((row) =>
      row.cells.map((each) => shapeOf(each.content).join("")),
    ),
  ).toEqual([["2", "B2"]]);
  expect(content.some((block) => block.type === "blockSdt")).toBe(false);
  publishCanonicalProjection({
    session,
    state: next,
    commit: session.prepareUndo(next).unwrap(),
  }).unwrap();
  expect(shapeOf(session.document.package.document.content)).toEqual(
    shapeOf(before.package.document.content),
  );
});

test("a pasted table drops the captures of a table under a control in its cell", () => {
  const inner: BlockContent = {
    type: "table",
    formatting: {
      sourceXml:
        '<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
      gridSourceXml:
        '<w:tblGrid xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
    },
    rows: [
      {
        type: "tableRow",
        cells: [{ type: "tableCell", content: [paragraph("N1", "1000D001")] }],
      },
    ],
  };
  const { session: source } = sessionOver([
    paragraph("Before", "10000001"),
    {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              content: [
                {
                  type: "blockSdt",
                  properties: { sdtType: "richText", id: 41 },
                  content: [inner, paragraph("", "1000D002")],
                },
                paragraph("A1", "1000D003"),
              ],
            },
          ],
        },
      ],
    },
  ]);
  const doc = source.projection.doc;
  const slice = doc.slice(doc.child(0).nodeSize, doc.content.size);
  const { session, state } = caretTarget();
  const prepared = prepareCanonicalPaste({
    session,
    state,
    slice,
    sourceDocument: readModel(source.document),
  });
  if (prepared.isErr()) throw prepared.error;
  publishCanonicalProjection({ session, state, commit: prepared.value });
  const tables: DeepReadonly<BlockContent>[] = [];
  const visit = (blocks: readonly DeepReadonly<BlockContent>[]) => {
    for (const block of blocks) {
      if (block.type === "table") {
        tables.push(block);
        for (const row of block.rows)
          for (const each of row.cells) visit(each.content);
      }
      if (block.type === "blockSdt") visit(block.content);
    }
  };
  visit(session.document.package.document.content);
  expect(tables).toHaveLength(2);
  for (const table of tables) {
    if (table.type !== "table") continue;
    expect(table.formatting?.sourceXml).toBeUndefined();
    expect(table.formatting?.gridSourceXml).toBeUndefined();
  }
});

const PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=",
  ),
  (char) => char.charCodeAt(0),
);

/** A body whose second paragraph holds an image of the package's own. */
const withImage = (): Document => ({
  package: {
    document: {
      content: [
        paragraph("Before", "10000001"),
        {
          type: "paragraph",
          paraId: "10000002",
          content: [
            {
              type: "run",
              content: [
                {
                  type: "drawing",
                  image: {
                    type: "image",
                    rId: relationshipIdOf("rId5") ?? panic("rId5"),
                    size: { width: 9525, height: 9525 },
                    wrap: { type: "inline" },
                  },
                },
              ],
            },
          ],
        },
        paragraph("After", "10000003"),
        paragraph("Tail", "10000004"),
      ],
    },
    relationships: new Map([
      [
        "rId5",
        {
          id: "rId5",
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          target: "media/image1.png",
        },
      ],
    ]),
    media: new Map([
      [
        "word/media/image1.png",
        {
          path: "word/media/image1.png",
          mimeType: "image/png",
          data: PNG.buffer,
        },
      ],
    ]),
  },
});

/** The relationship ids of the images the blocks hold, in order. */
const imageIds = (blocks: readonly DeepReadonly<BlockContent>[]): string[] =>
  JSON.stringify(blocks).match(/"rId":"[^"]*"/gu) ?? [];

describe.each([{ suggesting: false }, { suggesting: true }])(
  "a dragged image of the package's own (suggesting: $suggesting)",
  ({ suggesting }) => {
    test("moves under its relationship, undoes, and survives a save", async () => {
      const session = createCanonicalSession(withImage()).unwrap();
      session.setMode(
        suggesting
          ? { type: "suggesting", author: "Mover" }
          : { type: "editing" },
      );
      const doc = session.projection.doc;
      const before = session.document;
      // The image paragraph whole, from the end of "Before" to the start of
      // "After".
      const from = doc.child(0).nodeSize - 1;
      const to = doc.child(0).nodeSize + doc.child(1).nodeSize + 1;
      const slice = doc.slice(from, to);
      const state = EditorState.create({ doc });
      const prepared = prepareCanonicalPaste({
        session,
        state,
        slice,
        moveSource: { from, to },
        moveTarget: textStart(doc, "Tail") + 2,
      });
      if (prepared.isErr()) throw prepared.error;
      const next = publishCanonicalProjection({
        session,
        state,
        commit: prepared.value,
      }).unwrap().state;
      const moved = session.document;
      // No media or relationship was added: the image kept its own.
      expect([...(moved.package.relationships?.keys() ?? [])]).toEqual([
        "rId5",
      ]);
      expect([...(moved.package.media?.keys() ?? [])]).toEqual([
        "word/media/image1.png",
      ]);
      expect(imageIds(moved.package.document.content)).toEqual(
        suggesting ? ['"rId":"rId5"', '"rId":"rId5"'] : ['"rId":"rId5"'],
      );
      const reopened = await parseDocx(await createDocx(readModel(moved)), {
        preloadFonts: false,
      });
      expect(imageIds(reopened.package.document.content).length).toBe(
        imageIds(moved.package.document.content).length,
      );
      publishCanonicalProjection({
        session,
        state: next,
        commit: session.prepareUndo(next).unwrap(),
      }).unwrap();
      expect(shapeOf(session.document.package.document.content)).toEqual(
        shapeOf(before.package.document.content),
      );
      expect(imageIds(session.document.package.document.content)).toEqual([
        '"rId":"rId5"',
      ]);
    });
  },
);

test("dragged kept markup naming a relationship of the package moves, undoes, and survives a save", async () => {
  const xml =
    '<w:altChunk xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId7"/>';
  // A package whose body names its own part through the kept markup.
  const zip = await ZipArchive.read(
    await createDocx(
      documentOf([
        paragraph("Before", "10000001"),
        { type: "preservedBlock", xml },
        paragraph("After", "10000003"),
        paragraph("Tail", "10000004"),
      ]),
    ),
  );
  const rels = "word/_rels/document.xml.rels";
  const relsXml =
    (await zip.entry(rels)?.text()) ?? panic("The package has no rels.");
  zip.set(
    rels,
    new TextEncoder().encode(
      relsXml.replace(
        "</Relationships>",
        '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk.htm"/></Relationships>',
      ),
    ),
  );
  zip.set("word/afchunk.htm", new TextEncoder().encode("<p>chunk</p>"));
  const document = await parseDocx((await zip.generate()).buffer, {
    preloadFonts: false,
  });
  expect(document.package.relationships?.get("rId7")?.target).toBe(
    "afchunk.htm",
  );
  const session = createCanonicalSession(document).unwrap();
  const doc = session.projection.doc;
  const before = shapeOf(session.document.package.document.content);
  const from = doc.child(0).nodeSize - 1;
  const to = doc.child(0).nodeSize + doc.child(1).nodeSize + 1;
  const slice = doc.slice(from, to);
  const state = EditorState.create({ doc });
  const prepared = prepareCanonicalPaste({
    session,
    state,
    slice,
    moveSource: { from, to },
    moveTarget: textStart(doc, "Tail") + 2,
  });
  if (prepared.isErr()) throw prepared.error;
  const next = publishCanonicalProjection({
    session,
    state,
    commit: prepared.value,
  }).unwrap().state;
  const moved = shapeOf(session.document.package.document.content);
  expect(moved).toEqual(["BeforeAfter", "Ta", "preservedBlock", "il"]);
  const saved = await createDocx(readModel(session.document));
  const reopened = await parseDocx(saved, { preloadFonts: false });
  expect(shapeOf(reopened.package.document.content)).toEqual(moved);
  expect(reopened.package.relationships?.get("rId7")?.target).toBe(
    "afchunk.htm",
  );
  publishCanonicalProjection({
    session,
    state: next,
    commit: session.prepareUndo(next).unwrap(),
  }).unwrap();
  expect(shapeOf(session.document.package.document.content)).toEqual(before);
});

test("a dragged control holding a footnote reference moves with its note, undoes, and survives a save", async () => {
  const document = documentOf([
    paragraph("Before", "10000001"),
    {
      type: "blockSdt",
      properties: { sdtType: "richText", id: 41 },
      content: [
        {
          type: "paragraph",
          paraId: "1000A001",
          content: [
            { type: "run", content: [{ type: "text", text: "Noted" }] },
            { type: "run", content: [{ type: "footnoteRef", id: 1 }] },
          ],
        },
      ],
    },
    paragraph("After", "10000003"),
    paragraph("Tail", "10000004"),
  ]);
  document.package.footnotes = [
    { type: "footnote", id: 1, content: [paragraph("The note", "1000F001")] },
  ];
  const session = createCanonicalSession(document).unwrap();
  const doc = session.projection.doc;
  const before = session.document;
  const from = doc.child(0).nodeSize - 1;
  const to = doc.child(0).nodeSize + doc.child(1).nodeSize + 1;
  const slice = doc.slice(from, to);
  const state = EditorState.create({ doc });
  const prepared = prepareCanonicalPaste({
    session,
    state,
    slice,
    moveSource: { from, to },
    moveTarget: textStart(doc, "Tail") + 2,
  });
  if (prepared.isErr()) throw prepared.error;
  const next = publishCanonicalProjection({
    session,
    state,
    commit: prepared.value,
  }).unwrap().state;
  const moved = session.document;
  expect(shapeOf(moved.package.document.content)).toEqual([
    "BeforeAfter",
    "Ta",
    { blockSdt: ["Noted"] },
    "il",
  ]);
  const references = JSON.stringify(moved.package.document.content).match(
    /"footnoteRef","id":1/gu,
  );
  expect(references).toHaveLength(1);
  expect(shapeOf(moved.package.footnotes?.at(0)?.content ?? [])).toEqual([
    "The note",
  ]);
  const reopened = await parseDocx(await createDocx(readModel(moved)), {
    preloadFonts: false,
  });
  expect(shapeOf(reopened.package.document.content)).toEqual(
    shapeOf(moved.package.document.content),
  );
  expect(
    (reopened.package.footnotes ?? []).some(
      (note) => shapeOf(note.content).join("") === "The note",
    ),
  ).toBe(true);
  publishCanonicalProjection({
    session,
    state: next,
    commit: session.prepareUndo(next).unwrap(),
  }).unwrap();
  expect(session.document.package.footnotes).toEqual(before.package.footnotes);
  expect(shapeOf(session.document.package.document.content)).toEqual(
    shapeOf(before.package.document.content),
  );
});
