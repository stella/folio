/**
 * A `LISTNUM` field that opens a numbered paragraph is drawn as part of the
 * list marker. It stays in the paragraph's content as a capture of the markup
 * it was read from, showing nothing, so a save owes the file the field as it
 * was: the instruction, the field characters, the cached result with its
 * formatting, the tab after it, and where they sat.
 *
 * A capture is hidden only while it opens a paragraph whose marker shows it.
 * The save applies that rule to whatever the editor hands it: a capture an
 * edit took out of that position, or put under a tracked change, is written
 * as the ordinary field it stands for.
 *
 * The examples come first, one behaviour each. The properties after them
 * generate paragraphs and edits and read the result with oracles that share
 * nothing: the saved markup, read without the parser and compared with the
 * markup the paragraph was authored from, and the reopened document. The
 * last group pins what a document with no such field projects and saves to.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { createHash } from "node:crypto";
import { Window } from "happy-dom";
import { DOMParser, DOMSerializer, type Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { CLEARED_LIST_RENDERING_ATTRS } from "../prosemirror/listMarker";
import { paragraphNumberingAttr } from "../prosemirror/numberingAttr";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph } from "../types/document";
import {
  bodyParagraphs,
  cachedDisplay,
  contentShapes,
  documentXmlOf,
  editorState,
  expectedTokens,
  fieldResultsInFile,
  fieldResultsShown,
  type FieldSpec,
  foldedCaptureNodes,
  foldedFieldsOf,
  foldFaults,
  type GapMarker,
  inlineTokens,
  layoutMarkers,
  listNumberFieldDocx,
  MARKER_KINDS,
  modelMarkers,
  openDocx,
  paragraphMarkupOf,
  paragraphNode,
  type ParagraphSpec,
  PLAIN_PARAGRAPH_ID,
  positionOfText,
  type SavedDocx,
  saveDocx,
  typedInto,
  typeInto,
} from "./__tests__/listNumberFieldFixture";
import { repackDocx } from "./rezip";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { paragraphNumberingReference } from "@stll/docx-core/model";

setDefaultTimeout(propertyTestTimeout(30_000));

const LEADING: ParagraphSpec = {
  paraId: "20000001",
  marker: "decimal",
  fields: [
    {
      instruction: " LISTNUM  LegalDefault \\l 3 ",
      result: "(a)",
      formatting: "bold",
      before: "",
      gap: ["bookmark"],
      tab: true,
    },
  ],
  body: "Body text",
};

/** One field that opens the paragraph, and one behind its text. */
const INLINE: ParagraphSpec = {
  paraId: "20000002",
  marker: "decimal",
  fields: [
    {
      instruction: " LISTNUM ",
      result: "(b)",
      formatting: "plain",
      before: "",
      gap: [],
      tab: true,
    },
    {
      instruction: "LISTNUM \\l 4",
      result: "(i)",
      formatting: "plain",
      before: "first; ",
      gap: ["comment"],
      tab: false,
    },
  ],
  body: " second",
};

const SPECS = [LEADING, INLINE];

const SIMPLE: ParagraphSpec = {
  paraId: "20000001",
  marker: "decimal",
  fields: [
    {
      instruction: " LISTNUM ",
      result: "(a)",
      formatting: "plain",
      before: "",
      gap: [],
      tab: true,
    },
  ],
  body: "cd",
};

const SIMPLE_FIELD_TOKENS = [
  "fldChar:begin",
  "code: LISTNUM ",
  "fldChar:separate",
  "text:(a)",
  "fldChar:end",
  "tab",
];

const LEADING_FIELD_TOKENS = [
  "fldChar:begin",
  "code: LISTNUM  LegalDefault \\l 3 ",
  "fldChar:separate",
  "text:(a)",
  "fldChar:end",
  "bookmarkStart",
  "bookmarkEnd",
  "tab",
];

const ALL_CODES = ["code: LISTNUM  LegalDefault \\l 3 ", "code: LISTNUM ", "code:LISTNUM \\l 4"];

const codes = (tokens: readonly string[]): string[] =>
  tokens.filter((token) => token.startsWith("code:"));

const foldedKinds = (model: Document): string[][] =>
  contentShapes(model).map((shape) => shape.filter((kind) => kind.startsWith("folded:")));

const paragraphTokens = (model: Document): string[][] =>
  bodyParagraphs(model).map((paragraph) => inlineTokens(serializeParagraph(paragraph)));

/**
 * Every paragraph is in the form the fold allows and shows the field results
 * its markup holds: none hidden behind a marker that does not show it, none
 * twice.
 */
const expectShownIsWritten = (paragraphs: readonly Paragraph[]): void => {
  for (const paragraph of paragraphs) {
    expect(foldFaults(paragraph)).toEqual([]);
    expect(fieldResultsShown(paragraph)).toBe(
      fieldResultsInFile(inlineTokens(serializeParagraph(paragraph))),
    );
  }
};

/**
 * What a save of `doc` writes, and the document read back: each shows exactly
 * the fields the file holds, and the file holds `expectedCodes`.
 */
const expectHonest = async (
  doc: PMNode,
  parsed: Document,
  expectedCodes: readonly string[],
): Promise<Document> => {
  const rebuilt = fromProseDoc(doc, parsed);
  expectShownIsWritten(bodyParagraphs(rebuilt));
  expect(paragraphTokens(rebuilt).flatMap(codes)).toEqual([...expectedCodes]);

  const reopened = await openDocx(await repackDocx(rebuilt, { updateModifiedDate: false }));
  expectShownIsWritten(bodyParagraphs(reopened));
  expect(paragraphTokens(reopened).flatMap(codes)).toEqual([...expectedCodes]);
  expect(bodyParagraphs(reopened).map(fieldResultsShown)).toEqual(
    bodyParagraphs(rebuilt).map(fieldResultsShown),
  );
  return rebuilt;
};

const expectSavedFields = async (
  saved: SavedDocx,
  original: Document,
  bodies: Record<string, string> = {},
): Promise<void> => {
  const xml = await documentXmlOf(saved.bytes);
  for (const spec of SPECS) {
    const expected = expectedTokens(
      { ...spec, body: bodies[spec.paraId] ?? spec.body },
      saved.rewritten(spec.paraId),
    );
    expect(inlineTokens(paragraphMarkupOf(xml, spec.paraId))).toEqual(expected);
  }
  // The cached result keeps the run properties it was authored with.
  expect(paragraphMarkupOf(xml, LEADING.paraId)).toMatch(
    /<w:b\/>(?:(?!<\/w:r>).)*<w:t[^>]*>\(a\)<\/w:t>/u,
  );

  const reopened = await openDocx(saved.bytes);
  expect(foldedKinds(reopened)).toEqual(foldedKinds(original));
  expect(modelMarkers(reopened)).toEqual(modelMarkers(original));
  expect(layoutMarkers(reopened)).toEqual(layoutMarkers(original));
  expectShownIsWritten(bodyParagraphs(reopened));
};

type Fixture = {
  buffer: ArrayBuffer;
  parsed: Document;
  /** The editor as it runs, with every plugin a mounted editor holds. */
  state: EditorState;
  /** The same document with nothing watching it. */
  bare: EditorState;
};

const fixture = async (
  specs: readonly ParagraphSpec[] = SPECS,
  options: Parameters<typeof listNumberFieldDocx>[1] = {},
): Promise<Fixture> => {
  const buffer = await listNumberFieldDocx(specs, options);
  const parsed = await openDocx(buffer);
  return {
    buffer,
    parsed,
    state: editorState(parsed),
    bare: EditorState.create({ doc: toProseDoc(parsed) }),
  };
};

const rewritten = async (document: Document): Promise<SavedDocx> => ({
  bytes: await repackDocx(document, { updateModifiedDate: false }),
  rewritten: () => true,
});

describe("the reader's fold of LISTNUM fields", () => {
  test("hides the field that opens a paragraph behind its marker, and no other", async () => {
    const { parsed } = await fixture();
    const [leading, inline] = bodyParagraphs(parsed);
    if (!leading || !inline) {
      throw new Error("The fixture opens with two numbered paragraphs");
    }

    expect(leading.listRendering?.marker.endsWith("\t(a)")).toBe(true);
    expect(inline.listRendering?.marker.endsWith("\t(b)")).toBe(true);
    expect(expectedTokens(LEADING).slice(0, LEADING_FIELD_TOKENS.length)).toEqual(
      LEADING_FIELD_TOKENS,
    );

    const [leadingShape, inlineShape] = contentShapes(parsed);
    expect(leadingShape).toEqual([
      "folded:field",
      "bookmarkStart",
      "bookmarkEnd",
      "folded:tab",
      "run",
    ]);
    // The field behind "first; " is on the line, where the text is.
    expect(inlineShape?.slice(0, 6)).toEqual([
      "folded:field",
      "folded:tab",
      "run",
      "complexField",
      "commentRangeStart",
      "run",
    ]);
    expect(paragraphTokens(parsed).flatMap(codes)).toEqual(ALL_CODES);
    expectShownIsWritten(bodyParagraphs(parsed));
  });

  test("a field with no cached result folds nothing", async () => {
    const { parsed } = await fixture([
      {
        ...LEADING,
        fields: [
          {
            instruction: " LISTNUM ",
            result: "",
            formatting: "plain",
            before: "",
            gap: [],
            tab: true,
          },
        ],
      },
    ]);

    expect(contentShapes(parsed).at(0)).toEqual(["complexField", "run"]);
  });

  test("a bullet's marker hides no field", async () => {
    const { parsed } = await fixture([{ ...SIMPLE, marker: "symbol" }]);

    expect(contentShapes(parsed).at(0)).toEqual(["complexField", "run"]);
    expectShownIsWritten(bodyParagraphs(parsed));
  });

  test("the editor carries a capture as one hidden atom and nothing else new", async () => {
    const { bare } = await fixture([SIMPLE]);

    const captures = foldedCaptureNodes(bare.doc);

    expect(captures.map(({ node }) => Object.keys(node.attrs).toSorted())).toEqual([
      ["foldedListNumber", "level", "text", "xml"],
      ["foldedListNumber", "level", "text", "xml"],
    ]);
    expect(captures.map(({ node }) => node.attrs["text"])).toEqual(["", ""]);
  });
});

describe("a capture and the clipboard", () => {
  test("a copied paragraph carries no capture, and nothing parses back as one", async () => {
    const { bare } = await fixture([SIMPLE]);
    const { node } = paragraphNode(bare.doc, SIMPLE.paraId);
    expect(foldedCaptureNodes(node)).toHaveLength(2);
    const window = new Window();
    const document = window.document as unknown as globalThis.Document;

    // The paragraph's content, as the neighbouring DOM round trips serialize it.
    const host = document.createElement("div");
    host.append(DOMSerializer.fromSchema(schema).serializeFragment(node.content, { document }));

    expect(host.innerHTML).not.toContain("LISTNUM");
    expect(host.innerHTML).not.toContain("data-docx-preserved-xml");
    const parsed = DOMParser.fromSchema(schema).parse(host);
    const preserved: PMNode[] = [];
    parsed.descendants((child) => {
      if (child.type.name === "preservedXml") {
        preserved.push(child);
      }
    });
    expect(preserved).toEqual([]);
    // The text came through; only the hidden field and tab stayed behind.
    expect(parsed.textContent).toContain(SIMPLE.body);
  });
});

describe("saving a paragraph whose list marker holds LISTNUM fields", () => {
  test("an untouched paragraph keeps its fields when another one is edited", async () => {
    const { buffer, parsed, state } = await fixture();
    const edited = typeInto(state, PLAIN_PARAGRAPH_ID, "Plain.", "!");

    const saved = await saveDocx(fromProseDoc(edited.doc, parsed), buffer, [PLAIN_PARAGRAPH_ID]);

    await expectSavedFields(saved, parsed);
    expect(await documentXmlOf(saved.bytes)).toContain("P!lain.");
  });

  test("an edit elsewhere in the paragraph keeps its fields", async () => {
    const { buffer, parsed, state } = await fixture();
    const edited = typeInto(
      typeInto(state, LEADING.paraId, LEADING.body, "!"),
      INLINE.paraId,
      INLINE.body,
      "!",
    );

    const saved = await saveDocx(fromProseDoc(edited.doc, parsed), buffer, [
      LEADING.paraId,
      INLINE.paraId,
    ]);

    await expectSavedFields(saved, parsed, {
      [LEADING.paraId]: typedInto(LEADING.body, "!"),
      [INLINE.paraId]: typedInto(INLINE.body, "!"),
    });
  });

  test("a full rewrite keeps the fields of every paragraph", async () => {
    const { parsed, state } = await fixture();

    const saved = await rewritten(fromProseDoc(state.doc, parsed));

    await expectSavedFields(saved, parsed);
  });

  test("a second save writes what the first one did", async () => {
    const { parsed, bare } = await fixture();
    const once = await repackDocx(fromProseDoc(bare.doc, parsed), { updateModifiedDate: false });
    const reopened = await openDocx(once);

    const twice = await repackDocx(fromProseDoc(toProseDoc(reopened), reopened), {
      updateModifiedDate: false,
    });

    expect(await documentXmlOf(twice)).toBe(await documentXmlOf(once));
  });

  test("a field whose display is a recorded numbering change is written as it was read", async () => {
    const buffer = await listNumberFieldDocx([LEADING], {
      authored: {
        [LEADING.paraId]:
          `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
          `<w:r><w:instrText xml:space="preserve"> LISTNUM </w:instrText></w:r>` +
          `<w:r><w:fldChar w:fldCharType="end"><w:numberingChange w:id="9" w:author="A" w:date="2024-01-01T00:00:00Z" w:original="(a)"/></w:fldChar></w:r>` +
          `<w:r><w:tab/></w:r><w:r><w:t>Body text</w:t></w:r>`,
      },
    });
    const parsed = await openDocx(buffer);
    expect(bodyParagraphs(parsed).at(0)?.listRendering?.marker.endsWith("\t(a)")).toBe(true);
    const edited = typeInto(editorState(parsed), LEADING.paraId, "Body text", "!");

    const saved = await saveDocx(fromProseDoc(edited.doc, parsed), buffer, [LEADING.paraId]);

    const markup = paragraphMarkupOf(await documentXmlOf(saved.bytes), LEADING.paraId);
    expect(markup).toMatch(/<w:numberingChange\b[^>]*w:original="\(a\)"/u);
    // The display stays on the end character: no separator, no result run.
    expect(inlineTokens(markup)).toEqual([
      "fldChar:begin",
      "code: LISTNUM ",
      "fldChar:end",
      "tab",
      "text:B!ody text",
    ]);
    expect(modelMarkers(await openDocx(saved.bytes))).toEqual(modelMarkers(parsed));
  });

  test("an untouched paragraph keeps its captures byte for byte, and the range markers around them in order", async () => {
    // A bookmark and a comment range that open ahead of the field; the
    // bookmark closes between the field and its tab, the comment after the text.
    const authored =
      `<w:bookmarkStart w:id="5" w:name="around"/><w:commentRangeStart w:id="1"/>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> LISTNUM </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:rPr><w:b/></w:rPr><w:t>(a)</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `<w:bookmarkEnd w:id="5"/><w:r><w:tab/></w:r><w:r><w:t>Body text</w:t></w:r>` +
      `<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>`;
    const around: ParagraphSpec = {
      ...SIMPLE,
      fields: [{ ...SIMPLE.fields[0], gap: ["comment"] } as FieldSpec],
    };
    const buffer = await listNumberFieldDocx([around], { authored: { [around.paraId]: authored } });
    const parsed = await openDocx(buffer);
    const source = paragraphMarkupOf(await documentXmlOf(buffer), around.paraId);
    const captures = (bodyParagraphs(parsed).at(0)?.content ?? []).flatMap((item) =>
      item.type === "preservedInline" && item.foldedListNumber ? [item.xml] : [],
    );
    expect(captures).toHaveLength(2);
    expect(inlineTokens(source)).toEqual([
      "bookmarkStart",
      "commentRangeStart",
      ...SIMPLE_FIELD_TOKENS.slice(0, 5),
      "bookmarkEnd",
      "tab",
      "text:Body text",
      "commentRangeEnd",
      "commentReference",
    ]);
    const edited = typeInto(editorState(parsed), PLAIN_PARAGRAPH_ID, "Plain.", "!");
    const document = fromProseDoc(edited.doc, parsed);

    // Another paragraph edited: where the save patches, this one is its source bytes.
    const patched = await saveDocx(document, buffer, [PLAIN_PARAGRAPH_ID]);
    if (!patched.rewritten(around.paraId)) {
      expect(paragraphMarkupOf(await documentXmlOf(patched.bytes), around.paraId)).toBe(source);
    }

    // The whole part written anew: the captures are still what was read, and
    // every marker stands where it stood.
    const whole = paragraphMarkupOf(
      await documentXmlOf((await rewritten(document)).bytes),
      around.paraId,
    );
    for (const capture of captures) {
      expect(whole).toContain(capture);
    }
    expect(inlineTokens(whole)).toEqual(inlineTokens(source));
  });
});

describe("the save's rule for a field an edit moved", () => {
  test("text typed at the very start of the paragraph is written behind the field and its tab", async () => {
    const { buffer, parsed, state } = await fixture([SIMPLE]);
    const { position } = paragraphNode(state.doc, SIMPLE.paraId);

    const typed = state.apply(state.tr.insertText("X", position + 1));

    const saved = await saveDocx(fromProseDoc(typed.doc, parsed), buffer, [SIMPLE.paraId]);
    const markup = paragraphMarkupOf(await documentXmlOf(saved.bytes), SIMPLE.paraId);
    expect(inlineTokens(markup)).toEqual([...SIMPLE_FIELD_TOKENS, "text:Xcd"]);
    const reopened = await openDocx(saved.bytes);
    expect(modelMarkers(reopened)).toEqual(modelMarkers(parsed));
    expectShownIsWritten(bodyParagraphs(reopened));
  });

  test("text typed between the field and its tab leaves the field hidden and the tab on the line", async () => {
    const { parsed, state } = await fixture([SIMPLE]);
    const [field, tab] = foldedCaptureNodes(state.doc);
    if (!field || !tab) {
      throw new Error("The paragraph opens with a field capture and a tab capture");
    }

    const typed = state.apply(state.tr.insertText("X", tab.position));

    const rebuilt = await expectHonest(typed.doc, parsed, ["code: LISTNUM "]);
    expect(paragraphTokens(rebuilt).at(0)).toEqual([
      ...SIMPLE_FIELD_TOKENS.slice(0, 5),
      "text:X",
      "tab",
      "text:cd",
    ]);
    expect(foldedKinds(rebuilt).at(0)).toEqual(["folded:field"]);
  });

  test("text typed after the tab, and deleted there, leaves the field where it was", async () => {
    const { buffer, parsed, state } = await fixture([SIMPLE]);
    const typed = typeInto(state, SIMPLE.paraId, "cd", "!");
    const from = positionOfText(typed.doc, SIMPLE.paraId, "d");
    const edited = typed.apply(typed.tr.delete(from, from + 1));

    const saved = await saveDocx(fromProseDoc(edited.doc, parsed), buffer, [SIMPLE.paraId]);

    const markup = paragraphMarkupOf(await documentXmlOf(saved.bytes), SIMPLE.paraId);
    expect(inlineTokens(markup)).toEqual([...SIMPLE_FIELD_TOKENS, "text:c!"]);
  });

  test("a split at the very start leaves an empty paragraph above, and the field with its text", async () => {
    const { parsed, state } = await fixture([SIMPLE]);
    const { position } = paragraphNode(state.doc, SIMPLE.paraId);

    const split = state.apply(state.tr.split(position + 1));

    const rebuilt = await expectHonest(split.doc, parsed, ["code: LISTNUM "]);
    const [first, second] = paragraphTokens(rebuilt);
    expect(first).toEqual([]);
    expect(second).toEqual([...SIMPLE_FIELD_TOKENS, "text:cd"]);
    // The empty paragraph's marker does not show a field it does not hold.
    expect(modelMarkers(rebuilt).at(0)?.includes("\t")).toBe(false);
    expect(modelMarkers(rebuilt).at(1)?.endsWith("\t(a)")).toBe(true);
  });

  test("a split in the text leaves the field with the first half", async () => {
    const { parsed, state } = await fixture();
    const middle = positionOfText(state.doc, LEADING.paraId, " text");

    const split = state.apply(state.tr.split(middle));

    const rebuilt = await expectHonest(split.doc, parsed, ALL_CODES);
    const [first, second] = paragraphTokens(rebuilt);
    expect(first).toEqual([...LEADING_FIELD_TOKENS, "text:Body"]);
    expect(second).toEqual(["text: text"]);
    expect(modelMarkers(rebuilt).at(1)?.includes("\t")).toBe(false);
  });

  test("a paragraph joined behind a numbered one has its field written on the line", async () => {
    const { parsed, state } = await fixture();
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);

    const joined = state.apply(state.tr.join(position + node.nodeSize));

    const rebuilt = await expectHonest(joined.doc, parsed, ALL_CODES);
    expect(bodyParagraphs(rebuilt)).toHaveLength(2);
    // The first paragraph's marker still shows its own field, and only that one.
    expect(foldedKinds(rebuilt).at(0)).toEqual(["folded:field", "folded:tab"]);
    expect(modelMarkers(rebuilt).at(0)?.endsWith("\t(a)")).toBe(true);
    const tokens = paragraphTokens(rebuilt).at(0) ?? [];
    expect(tokens.slice(0, LEADING_FIELD_TOKENS.length + 2)).toEqual([
      ...LEADING_FIELD_TOKENS,
      "text:Body text",
      "fldChar:begin",
    ]);
    expect(bodyParagraphs(rebuilt).slice(0, 1).map(fieldResultsShown)).toEqual(["(a) (b) (i)"]);
  });

  test("a paragraph taken out of its list has its field written on the line", async () => {
    const { parsed, state } = await fixture();
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);

    const unnumbered = state.apply(
      state.tr.setNodeMarkup(position, undefined, {
        ...node.attrs,
        ...CLEARED_LIST_RENDERING_ATTRS,
        numPr: null,
      }),
    );

    const rebuilt = await expectHonest(unnumbered.doc, parsed, ALL_CODES);
    expect(paragraphTokens(rebuilt).at(0)).toEqual([...LEADING_FIELD_TOKENS, "text:Body text"]);
    expect(foldedKinds(rebuilt).at(0)).toEqual([]);
    expect(bodyParagraphs(rebuilt).slice(0, 1).map(fieldResultsShown)).toEqual(["(a)"]);
  });

  test("a paragraph moved to another level, whose marker is drawn anew, has its field written on the line", async () => {
    const { parsed, state } = await fixture();
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);

    const moved = state.apply(
      state.tr.setNodeMarkup(position, undefined, {
        ...node.attrs,
        numPr: paragraphNumberingAttr(paragraphNumberingReference({ numId: 1, ilvl: 2 })),
        listMarker: "(%3)",
        listMarkerTemplate: "(%3)",
      }),
    );

    const rebuilt = await expectHonest(moved.doc, parsed, ALL_CODES);
    expect(paragraphTokens(rebuilt).at(0)).toEqual([...LEADING_FIELD_TOKENS, "text:Body text"]);
    expect(foldedKinds(rebuilt).at(0)).toEqual([]);
  });

  test("a numbered paragraph joined behind a plain one has its field written on the line", async () => {
    const { parsed, state } = await fixture([SIMPLE], { plainFirst: true });
    const { node, position } = paragraphNode(state.doc, PLAIN_PARAGRAPH_ID);

    const joined = state.apply(state.tr.join(position + node.nodeSize));

    const rebuilt = await expectHonest(joined.doc, parsed, ["code: LISTNUM "]);
    expect(paragraphTokens(rebuilt)).toEqual([["text:Plain.", ...SIMPLE_FIELD_TOKENS, "text:cd"]]);
    expect(bodyParagraphs(rebuilt).map(fieldResultsShown)).toEqual(["(a)"]);
  });

  test("a capture marked as deleted is written as a deleted field", async () => {
    const { parsed, state } = await fixture([SIMPLE]);
    const [field] = foldedCaptureNodes(state.doc);
    const deletion = state.schema.marks["deletion"]?.create({
      revisionId: 7,
      author: "Reviewer",
      date: "2026-01-01T00:00:00Z",
    });
    if (!field || !deletion) {
      throw new Error("The paragraph opens with a field capture, and the schema tracks deletions");
    }

    const marked = state.apply(state.tr.addMark(field.position, field.position + 1, deletion));

    const [paragraph] = bodyParagraphs(fromProseDoc(marked.doc, parsed));
    if (!paragraph) {
      throw new Error("The document keeps its first paragraph");
    }
    expect(foldFaults(paragraph)).toEqual([]);
    const xml = serializeParagraph(paragraph);
    expect(xml).toContain("<w:delInstrText");
    expect(xml).not.toContain("<w:instrText");
    expect(xml).toMatch(/<w:delText[^>]*>\(a\)<\/w:delText>/u);
  });

  test("a field deleted with the content around it is not written, and its marker stops showing it", async () => {
    const { parsed, state } = await fixture();
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);

    const emptied = state.apply(state.tr.delete(position + 1, position + node.nodeSize - 1));

    const rebuilt = await expectHonest(emptied.doc, parsed, ALL_CODES.slice(1));
    expect(paragraphTokens(rebuilt).at(0)).toEqual([]);
    expect(modelMarkers(rebuilt).at(0)?.includes("\t")).toBe(false);
  });
});

const GAP_MARKERS = [
  "bookmark",
  "bookmarkStart",
  "comment",
] as const satisfies readonly GapMarker[];
const RESULT_FORMATTING = ["plain", "bold", "symbol"] as const;

const fieldArbitrary: fc.Arbitrary<FieldSpec> = fc.record({
  instruction: fc.constantFrom(
    " LISTNUM ",
    "LISTNUM",
    " LISTNUM  LegalDefault \\l 3 ",
    " listnum NumberDefault \\s 2 ",
  ),
  // A literal percent sign among them; the fixed example below adds a tab
  // inside the cached result.
  result: fc.constantFrom("(a)", "(ii)", "50%", "%"),
  formatting: fc.constantFrom(...RESULT_FORMATTING),
  before: fc.constantFrom("", "", "and ", "x"),
  gap: fc.uniqueArray(fc.constantFrom(...GAP_MARKERS), { maxLength: 3 }),
  tab: fc.boolean(),
});

/**
 * A comment opens after the last field only, as the last marker ahead of its
 * tab: its range runs to the paragraph's end, and one range per paragraph is
 * all the fixture's comments part describes.
 */
const foldable = (fields: FieldSpec[]): FieldSpec[] =>
  fields.map((field, index): FieldSpec => {
    const last = index === fields.length - 1;
    const bookmarks = field.gap.filter((marker) => marker !== "comment");
    const commented = last && field.gap.includes("comment");
    return { ...field, gap: commented ? [...bookmarks, "comment"] : bookmarks };
  });

const paragraphsArbitrary: fc.Arbitrary<ParagraphSpec[]> = fc
  .array(
    fc.record({
      marker: fc.constantFrom(...MARKER_KINDS),
      fields: fc.array(fieldArbitrary, { minLength: 1, maxLength: 3 }).map(foldable),
      body: fc.constantFrom("Body", "Tail text", "100% of it"),
    }),
    { minLength: 1, maxLength: 2 },
  )
  .map((paragraphs) =>
    paragraphs.map((paragraph, index) => ({ ...paragraph, paraId: `2000000${index + 1}` })),
  );

const TYPED = "!";

const expectSaved = async (
  saved: SavedDocx,
  original: Document,
  specs: readonly ParagraphSpec[],
): Promise<void> => {
  const xml = await documentXmlOf(saved.bytes);
  for (const spec of specs) {
    expect(inlineTokens(paragraphMarkupOf(xml, spec.paraId))).toEqual(
      expectedTokens(spec, saved.rewritten(spec.paraId)),
    );
  }

  const reopened = await openDocx(saved.bytes);
  expect(foldedKinds(reopened)).toEqual(foldedKinds(original));
  expect(modelMarkers(reopened)).toEqual(modelMarkers(original));
  expect(layoutMarkers(reopened)).toEqual(layoutMarkers(original));
  expectShownIsWritten(bodyParagraphs(reopened));
};

/** What the reader owes each paragraph before any save is looked at. */
const expectFolded = (parsed: Document, specs: readonly ParagraphSpec[]): void => {
  const paragraphs = bodyParagraphs(parsed);
  for (const [index, spec] of specs.entries()) {
    const paragraph = paragraphs.at(index);
    if (!paragraph) {
      throw new Error(`The document has no paragraph ${spec.paraId}`);
    }
    // Only the fields that open the paragraph are behind the marker.
    const folded = foldedFieldsOf(spec);
    expect(foldedKinds(parsed).at(index)).toEqual(
      folded.flatMap(({ tab }) => (tab ? ["folded:field", "folded:tab"] : ["folded:field"])),
    );
    const marker = paragraph.listRendering?.marker ?? "";
    if (folded.length > 0) {
      expect(marker.endsWith(`\t${cachedDisplay(folded)}`)).toBe(true);
    } else if (spec.marker !== "symbol") {
      expect(marker.includes("\t")).toBe(false);
    }
    expect(fieldResultsShown(paragraph)).toBe(cachedDisplay(spec.fields));
  }
};

describe("saving generated paragraphs whose list markers hold LISTNUM fields", () => {
  test(
    "every field goes back where it stood and the markers stay as they were",
    async () => {
      await assertProperty(
        fc.asyncProperty(paragraphsArbitrary, async (specs) => {
          const buffer = await listNumberFieldDocx(specs);
          const parsed = await openDocx(buffer);
          expectFolded(parsed, specs);
          const opened = editorState(parsed);

          // No paragraph of its own edited.
          const beside = typeInto(opened, PLAIN_PARAGRAPH_ID, "Plain.", TYPED);
          await expectSaved(
            await saveDocx(fromProseDoc(beside.doc, parsed), buffer, [PLAIN_PARAGRAPH_ID]),
            parsed,
            specs,
          );

          // Text typed into every paragraph, after its fields.
          let edited = opened;
          for (const spec of specs) {
            edited = typeInto(edited, spec.paraId, spec.body, TYPED);
          }
          await expectSaved(
            await saveDocx(
              fromProseDoc(edited.doc, parsed),
              buffer,
              specs.map(({ paraId }) => paraId),
            ),
            parsed,
            specs.map((spec) => ({ ...spec, body: typedInto(spec.body, TYPED) })),
          );

          // The whole part rewritten.
          await expectSaved(await rewritten(fromProseDoc(opened.doc, parsed)), parsed, specs);
        }),
        {
          numRuns: 30,
          examples: [
            // A comment that opens between a field on the line and its tab.
            [
              [
                {
                  paraId: "20000001",
                  marker: "decimal",
                  fields: [
                    {
                      instruction: " LISTNUM ",
                      result: "(a)",
                      formatting: "plain",
                      before: "and ",
                      gap: ["comment"],
                      tab: true,
                    },
                  ],
                  body: "Body",
                },
              ],
            ],
            [
              [
                {
                  paraId: "20000001",
                  marker: "percent",
                  fields: [
                    {
                      instruction: " LISTNUM ",
                      result: "a\tb",
                      formatting: "symbol",
                      before: "",
                      gap: ["bookmark"],
                      tab: true,
                    },
                    {
                      instruction: "LISTNUM",
                      result: "50%",
                      formatting: "bold",
                      before: "",
                      gap: ["bookmarkStart"],
                      tab: true,
                    },
                    {
                      instruction: " LISTNUM  LegalDefault \\l 3 ",
                      result: "(a)",
                      formatting: "symbol",
                      before: "x",
                      gap: ["comment"],
                      tab: false,
                    },
                  ],
                  body: "Body",
                },
                {
                  paraId: "20000002",
                  marker: "symbol",
                  fields: [
                    {
                      instruction: " listnum NumberDefault \\s 2 ",
                      result: "(ii)",
                      formatting: "plain",
                      before: "",
                      gap: [],
                      tab: true,
                    },
                  ],
                  body: "Tail text",
                },
              ],
            ],
          ],
        },
      );
    },
    propertyTestTimeout(180_000),
  ); // Each case zips a package, then saves and reopens it three times.
});

const EDITED: ParagraphSpec[] = [
  {
    paraId: "20000001",
    marker: "decimal",
    fields: [
      {
        instruction: " LISTNUM ",
        result: "(a)",
        formatting: "plain",
        before: "",
        gap: [],
        tab: true,
      },
      {
        instruction: "LISTNUM",
        result: "(b)",
        formatting: "bold",
        before: "one ",
        gap: [],
        tab: false,
      },
    ],
    body: "first body",
  },
  {
    paraId: "20000002",
    marker: "decimal",
    fields: [
      {
        instruction: " LISTNUM ",
        result: "(c)",
        formatting: "plain",
        before: "",
        gap: [],
        tab: true,
      },
    ],
    body: "second body",
  },
  {
    paraId: "20000003",
    marker: "percent",
    fields: [
      {
        instruction: " LISTNUM ",
        result: "(d)",
        formatting: "plain",
        before: "two ",
        gap: [],
        tab: true,
      },
    ],
    body: "third body",
  },
];

const RESULTS = ["text:(a)", "text:(b)", "text:(c)", "text:(d)"];

type Place = { paragraph: number; offset: number };

/** A place in the document: a paragraph, and how far into its content. */
const placeArbitrary: fc.Arbitrary<Place> = fc.record({
  paragraph: fc.nat({ max: 999 }).map((thousandths) => thousandths / 1000),
  offset: fc.nat({ max: 1000 }).map((thousandths) => thousandths / 1000),
});

type Step =
  | { kind: "split" | "join" | "type"; at: Place }
  | { kind: "copy"; from: Place; to: Place; at: Place };

const stepArbitrary: fc.Arbitrary<Step> = fc.oneof(
  fc.record({ kind: fc.constant("split" as const), at: placeArbitrary }),
  fc.record({ kind: fc.constant("join" as const), at: placeArbitrary }),
  fc.record({ kind: fc.constant("type" as const), at: placeArbitrary }),
  fc.record({
    kind: fc.constant("copy" as const),
    from: placeArbitrary,
    to: placeArbitrary,
    at: placeArbitrary,
  }),
);

const topParagraphs = (doc: PMNode): { node: PMNode; position: number }[] => {
  const paragraphs: { node: PMNode; position: number }[] = [];
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  doc.forEach((node, position) => {
    if (node.type.name === "paragraph") {
      paragraphs.push({ node, position });
    }
  });
  return paragraphs;
};

/** The document position `place` names, inside a paragraph's content. */
const positionOf = (doc: PMNode, place: Place): number => {
  const paragraphs = topParagraphs(doc);
  const paragraph = paragraphs[Math.floor(place.paragraph * paragraphs.length)];
  if (!paragraph) {
    throw new Error("The document holds no paragraph");
  }
  return paragraph.position + 1 + Math.round(place.offset * paragraph.node.content.size);
};

const applyStep = (state: EditorState, step: Step): EditorState => {
  switch (step.kind) {
    case "split":
      return state.apply(state.tr.split(positionOf(state.doc, step.at)));
    case "join": {
      const paragraphs = topParagraphs(state.doc);
      const first = paragraphs[Math.floor(step.at.paragraph * (paragraphs.length - 1))];
      // One paragraph has nothing to join.
      if (!first || paragraphs.length < 2) {
        return state;
      }
      return state.apply(state.tr.join(first.position + first.node.nodeSize));
    }
    case "type":
      return state.apply(state.tr.insertText("x", positionOf(state.doc, step.at)));
    case "copy": {
      const a = positionOf(state.doc, step.from);
      const b = positionOf(state.doc, step.to);
      const copied = state.doc.slice(Math.min(a, b), Math.max(a, b));
      const at = positionOf(state.doc, step.at);
      // The copy takes its captures along, as content does.
      return state.apply(state.tr.replaceRange(at, at, copied));
    }
    default: {
      const unhandled: never = step;
      throw new Error(`Unhandled step ${JSON.stringify(unhandled)}`);
    }
  }
};

describe("saving after splits, joins, typing and copies in the editor", () => {
  test(
    "no field is written hidden unless its paragraph's marker shows it, and none is lost",
    async () => {
      const parsed = await openDocx(await listNumberFieldDocx(EDITED));
      const opened = editorState(parsed);
      // The field behind "one " and the field behind "two " are on the line.
      expect(foldedCaptureNodes(opened.doc)).toHaveLength(4);

      assertProperty(
        fc.property(fc.array(stepArbitrary, { minLength: 1, maxLength: 12 }), (steps) => {
          let state = opened;
          for (const step of steps) {
            state = applyStep(state, step);
          }

          const paragraphs = bodyParagraphs(fromProseDoc(state.doc, parsed));
          expectShownIsWritten(paragraphs);
          const written = paragraphs.flatMap((paragraph) =>
            inlineTokens(serializeParagraph(paragraph)),
          );
          // Nothing here deletes, so every field is still written at least once.
          for (const result of RESULTS) {
            expect(written).toContain(result);
          }
        }),
        {
          numRuns: 150,
          examples: [
            // Text typed at the very start of a paragraph whose marker shows a field.
            [[{ kind: "type", at: { paragraph: 0, offset: 0 } }]],
            // A whole paragraph copied and put ahead of itself.
            [
              [
                {
                  kind: "copy",
                  from: { paragraph: 0, offset: 0 },
                  to: { paragraph: 0.25, offset: 0 },
                  at: { paragraph: 0, offset: 0 },
                },
              ],
            ],
          ],
        },
      );
    },
    propertyTestTimeout(120_000),
  );
});

type Pinned = {
  path: string;
  /** The projected document's JSON: its SHA-256, or only its length where ids are minted per process. */
  projection: { digest: string; length: number } | { length: number };
  /**
   * The saved `word/document.xml`: the file that holds it, compared as text so
   * a difference shows where it is, or its SHA-256 and length where the markup
   * is too long to keep beside the fixture.
   */
  saved: { file: string } | { digest: string; length: number };
};

/**
 * Taken with the reader, the projection and the save as they stood before the
 * fold kept its fields in the paragraph content. `docx-editor-demo.docx` is
 * the fixture with numbered paragraphs, the ones the fold looks at.
 */
const FIXTURES: readonly Pinned[] = [
  {
    path: "tests/visual/fixtures/sample.docx",
    projection: {
      digest: "2213484f1e631a4a2e3223ff8ec4fca38e5eb1ee4dea358fdca9b832aee3d04b",
      length: 90_296,
    },
    saved: { file: "sample.document.xml" },
  },
  {
    path: "packages/core/src/docx/__tests__/__fixtures__/corpus/step3-header-footer-fields.docx",
    projection: {
      digest: "a68195ab6c8e32a6df0e09ed04cc6131467dd2f5035fda94da7580149e7726aa",
      length: 8803,
    },
    saved: { file: "step3-header-footer-fields.document.xml" },
  },
  {
    path: "packages/core/src/docx/__tests__/__fixtures__/corpus/step3-footnotes.docx",
    projection: {
      digest: "db9ec37655bc5ed3c05a45c698f76a09f36fa5786d43bd843746ddcadc844c22",
      length: 4271,
    },
    saved: { file: "step3-footnotes.document.xml" },
  },
  {
    path: "tests/visual/fixtures/docx-editor-demo.docx",
    projection: { length: 445_121 },
    saved: {
      digest: "6bac5af14f8a58ea6627768ea747a8fc8ec69ce3e3edae1c41d720de84f81dee",
      length: 59_324,
    },
  },
];

const sha256 = (text: string): string => createHash("sha256").update(text).digest("hex");

describe("a document with no LISTNUM field", () => {
  for (const { path, projection, saved } of FIXTURES) {
    test(`${path} projects and saves as it did`, async () => {
      const buffer = await Bun.file(new URL(`../../../../${path}`, import.meta.url)).arrayBuffer();
      expect(/LISTNUM/iu.test(await documentXmlOf(buffer))).toBe(false);

      const parsed = await openDocx(buffer);
      const doc = toProseDoc(parsed);
      const json = JSON.stringify(doc.toJSON());

      expect(json).not.toContain("foldedListNumber");
      // The fixture and the length go with the digest, so a mismatch says
      // which document moved and whether it grew or shrank.
      expect({
        fixture: path,
        length: json.length,
        ...("digest" in projection ? { digest: sha256(json) } : {}),
      }).toEqual({ fixture: path, ...projection });

      const written = await documentXmlOf(
        await repackDocx(fromProseDoc(doc, parsed), { updateModifiedDate: false }),
      );
      if ("file" in saved) {
        const reference = await Bun.file(
          new URL(`./__tests__/__fixtures__/list-number-invisible/${saved.file}`, import.meta.url),
        ).text();
        expect(written).toBe(reference);
      } else {
        expect({ fixture: path, length: written.length, digest: sha256(written) }).toEqual({
          fixture: path,
          ...saved,
        });
      }
    });
  }
});
