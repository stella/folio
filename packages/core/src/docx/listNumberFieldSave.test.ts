/**
 * A `LISTNUM` field that opens a numbered paragraph is drawn as part of the
 * list marker. It stays in the paragraph's content as a capture of the markup
 * it was read from, showing nothing, so a save owes the file the field as it
 * was: the instruction, the field characters, the cached result with its
 * formatting, the tab after it, and where they sat.
 *
 * A capture is hidden only while its paragraph's marker shows it. Wherever an
 * edit takes it out of that, it becomes the ordinary field again, on the
 * line, so the page never hides a field the file holds.
 */

import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import {
  foldedListNumberPlugin,
  unfoldPastedListNumberFields,
} from "../prosemirror/foldedListNumber";
import { CLEARED_LIST_RENDERING_ATTRS } from "../prosemirror/listMarker";
import { paragraphNumberingAttr } from "../prosemirror/numberingAttr";
import type { Document } from "../types/document";
import {
  bodyParagraphs,
  contentShapes,
  documentXmlOf,
  expectedTokens,
  fieldResultsInFile,
  fieldResultsShown,
  foldedCaptureNodes,
  inlineTokens,
  layoutMarkers,
  listNumberFieldDocx,
  liveFoldFaults,
  modelMarkers,
  openDocx,
  paragraphMarkupOf,
  paragraphNode,
  type ParagraphSpec,
  PLAIN_PARAGRAPH_ID,
  positionOfText,
  saveDocx,
  typedInto,
  typeInto,
  withSettledTail,
} from "./__tests__/listNumberFieldFixture";
import { repackDocx } from "./rezip";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { paragraphNumberingReference } from "@stll/docx-core/model";

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

/** No range markers: a paragraph that can be copied without copying an id. */
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

const FRESH_PARAGRAPH_ID = "2000000A";

/** Split at `position`, the new paragraph taking an id of its own as the editor gives it one. */
const splitAt = (state: EditorState, position: number): EditorState =>
  state.apply(
    state.tr.split(position).setNodeAttribute(position + 1, "paraId", FRESH_PARAGRAPH_ID),
  );

const codes = (tokens: readonly string[]): string[] =>
  tokens.filter((token) => token.startsWith("code:"));

const foldedKinds = (model: Document): string[][] =>
  contentShapes(model).map((shape) => shape.filter((kind) => kind.startsWith("folded:")));

const paragraphTokens = (model: Document): string[][] =>
  bodyParagraphs(model).map((paragraph) => inlineTokens(serializeParagraph(paragraph)));

/** Every paragraph shows the field results its markup holds: none hidden, none twice. */
const expectShownIsWritten = (model: Document): void => {
  for (const paragraph of bodyParagraphs(model)) {
    expect(fieldResultsShown(paragraph)).toBe(
      fieldResultsInFile(inlineTokens(serializeParagraph(paragraph))),
    );
  }
};

/**
 * The editor state, what a save of it writes, and the document read back:
 * each shows exactly the fields the file holds, and the file holds `codes`.
 */
const expectHonest = async (
  doc: PMNode,
  parsed: Document,
  expectedCodes: readonly string[],
): Promise<Document> => {
  expect(liveFoldFaults(doc)).toEqual([]);
  const rebuilt = fromProseDoc(doc, parsed);
  expectShownIsWritten(rebuilt);
  expect(paragraphTokens(rebuilt).flatMap(codes)).toEqual([...expectedCodes]);

  const reopened = await openDocx(await repackDocx(rebuilt, { updateModifiedDate: false }));
  expectShownIsWritten(reopened);
  expect(paragraphTokens(reopened).flatMap(codes)).toEqual([...expectedCodes]);
  expect(bodyParagraphs(reopened).map(fieldResultsShown)).toEqual(
    bodyParagraphs(rebuilt).map(fieldResultsShown),
  );
  return rebuilt;
};

const expectSavedFields = async (
  saved: ArrayBuffer,
  original: Document,
  bodies: Record<string, string> = {},
): Promise<void> => {
  const xml = await documentXmlOf(saved);
  for (const spec of SPECS) {
    const expected = expectedTokens({ ...spec, body: bodies[spec.paraId] ?? spec.body });
    expect(withSettledTail(inlineTokens(paragraphMarkupOf(xml, spec.paraId)))).toEqual(
      withSettledTail(expected),
    );
  }
  // The cached result keeps the run properties it was authored with.
  expect(paragraphMarkupOf(xml, LEADING.paraId)).toMatch(
    /<w:b\/>(?:(?!<\/w:r>).)*<w:t[^>]*>\(a\)<\/w:t>/u,
  );

  const reopened = await openDocx(saved);
  expect(foldedKinds(reopened)).toEqual(foldedKinds(original));
  expect(modelMarkers(reopened)).toEqual(modelMarkers(original));
  expect(layoutMarkers(reopened)).toEqual(layoutMarkers(original));
  expectShownIsWritten(reopened);
};

type Fixture = {
  buffer: ArrayBuffer;
  parsed: Document;
  /** The editor as it runs: every change is followed by the fold's own pass. */
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
  const doc = toProseDoc(parsed);
  return {
    buffer,
    parsed,
    state: EditorState.create({ doc, plugins: [foldedListNumberPlugin()] }),
    bare: EditorState.create({ doc }),
  };
};

describe("the reader's fold of LISTNUM fields", () => {
  test("hides the field that opens a paragraph behind its marker, and no other", async () => {
    const { parsed, state } = await fixture();
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
    expectShownIsWritten(parsed);
    expect(liveFoldFaults(state.doc)).toEqual([]);
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
    expectShownIsWritten(parsed);
  });
});

describe("saving a paragraph whose list marker holds LISTNUM fields", () => {
  test("an untouched paragraph keeps its fields when another one is edited", async () => {
    const { buffer, parsed, state } = await fixture();
    const edited = typeInto(state, PLAIN_PARAGRAPH_ID, "Plain.", "!");

    const saved = await saveDocx(fromProseDoc(edited.doc, parsed), buffer, [PLAIN_PARAGRAPH_ID]);

    await expectSavedFields(saved, parsed);
    expect(await documentXmlOf(saved)).toContain("P!lain.");
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

    const saved = await repackDocx(fromProseDoc(state.doc, parsed), { updateModifiedDate: false });

    await expectSavedFields(saved, parsed);
  });

  test("a second save writes what the first one did", async () => {
    const { parsed, state } = await fixture();
    const once = await repackDocx(fromProseDoc(state.doc, parsed), { updateModifiedDate: false });
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
    const edited = typeInto(
      EditorState.create({ doc: toProseDoc(parsed), plugins: [foldedListNumberPlugin()] }),
      LEADING.paraId,
      "Body text",
      "!",
    );

    const saved = await saveDocx(fromProseDoc(edited.doc, parsed), buffer, [LEADING.paraId]);

    const markup = paragraphMarkupOf(await documentXmlOf(saved), LEADING.paraId);
    expect(markup).toMatch(/<w:numberingChange\b[^>]*w:original="\(a\)"/u);
    // The display stays on the end character: no separator, no result run.
    expect(inlineTokens(markup)).toEqual([
      "fldChar:begin",
      "code: LISTNUM ",
      "fldChar:end",
      "tab",
      "text:B!ody text",
    ]);
    expect(modelMarkers(await openDocx(saved))).toEqual(modelMarkers(parsed));
  });
});

describe("typing around the field a marker shows", () => {
  for (const watched of [true, false]) {
    const editor = watched ? "the editor's own pass" : "the save alone";

    test(`text typed at the very start of the paragraph lands behind the field and its tab (${editor})`, async () => {
      const { buffer, parsed, state, bare } = await fixture([SIMPLE]);
      const start = watched ? state : bare;
      const { position } = paragraphNode(start.doc, SIMPLE.paraId);

      const typed = start.apply(start.tr.insertText("X", position + 1));

      if (watched) {
        expect(liveFoldFaults(typed.doc)).toEqual([]);
        expect(paragraphNode(typed.doc, SIMPLE.paraId).node.firstChild?.type.name).toBe(
          "preservedXml",
        );
      }
      const saved = await saveDocx(fromProseDoc(typed.doc, parsed), buffer, [SIMPLE.paraId]);
      const markup = paragraphMarkupOf(await documentXmlOf(saved), SIMPLE.paraId);
      expect(inlineTokens(markup)).toEqual([...SIMPLE_FIELD_TOKENS, "text:Xcd"]);
      const reopened = await openDocx(saved);
      expect(modelMarkers(reopened)).toEqual(modelMarkers(parsed));
      expectShownIsWritten(reopened);
    });
  }

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

    const markup = paragraphMarkupOf(await documentXmlOf(saved), SIMPLE.paraId);
    expect(inlineTokens(markup)).toEqual([...SIMPLE_FIELD_TOKENS, "text:c!"]);
    expect(liveFoldFaults(edited.doc)).toEqual([]);
  });

  test("Enter at the very start leaves an empty paragraph above, and the field with its text", async () => {
    const { parsed, state } = await fixture([SIMPLE]);
    const { position } = paragraphNode(state.doc, SIMPLE.paraId);

    const split = splitAt(state, position + 1);

    const rebuilt = await expectHonest(split.doc, parsed, ["code: LISTNUM "]);
    const [first, second] = paragraphTokens(rebuilt);
    expect(first).toEqual([]);
    expect(second).toEqual([...SIMPLE_FIELD_TOKENS, "text:cd"]);
    // The empty paragraph's marker no longer shows a field it does not hold.
    expect(modelMarkers(rebuilt).at(0)?.includes("\t")).toBe(false);
    expect(modelMarkers(rebuilt).at(1)?.endsWith("\t(a)")).toBe(true);
  });

  test("Enter in the text leaves the field with the first half", async () => {
    const { parsed, state } = await fixture();
    const middle = positionOfText(state.doc, LEADING.paraId, " text");

    const split = splitAt(state, middle);

    const rebuilt = await expectHonest(split.doc, parsed, ALL_CODES);
    const [first, second] = paragraphTokens(rebuilt);
    expect(first).toEqual([...LEADING_FIELD_TOKENS, "text:Body"]);
    expect(second).toEqual(["text: text"]);
    expect(modelMarkers(rebuilt).at(1)?.includes("\t")).toBe(false);
  });

  test("Backspace at the very start joins the paragraph behind the one before, with its field on the line", async () => {
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
    expect(
      fieldResultsShown(bodyParagraphs(rebuilt)[0] ?? { type: "paragraph", content: [] }),
    ).toBe("(a) (b) (i)");
  });
});

describe("a field its marker stops showing goes back on the line", () => {
  test("a paragraph taken out of its list", async () => {
    const { parsed, state } = await fixture();
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);

    const unnumbered = state.apply(
      state.tr.setNodeMarkup(position, undefined, {
        ...node.attrs,
        ...CLEARED_LIST_RENDERING_ATTRS,
        numPr: null,
      }),
    );

    expect(foldedCaptureNodes(paragraphNode(unnumbered.doc, LEADING.paraId).node)).toHaveLength(0);
    const rebuilt = await expectHonest(unnumbered.doc, parsed, ALL_CODES);
    expect(paragraphTokens(rebuilt).at(0)).toEqual([...LEADING_FIELD_TOKENS, "text:Body text"]);
    expect(foldedKinds(rebuilt).at(0)).toEqual([]);
    expect(
      fieldResultsShown(bodyParagraphs(rebuilt)[0] ?? { type: "paragraph", content: [] }),
    ).toBe("(a)");
  });

  test("a paragraph moved to another level, whose marker is drawn anew", async () => {
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

    expect(foldedCaptureNodes(paragraphNode(moved.doc, LEADING.paraId).node)).toHaveLength(0);
    const rebuilt = await expectHonest(moved.doc, parsed, ALL_CODES);
    expect(paragraphTokens(rebuilt).at(0)).toEqual([...LEADING_FIELD_TOKENS, "text:Body text"]);
  });

  test("a numbered paragraph joined behind a plain one", async () => {
    const { parsed, state } = await fixture([SIMPLE], { plainFirst: true });
    const { node, position } = paragraphNode(state.doc, PLAIN_PARAGRAPH_ID);

    const joined = state.apply(state.tr.join(position + node.nodeSize));

    expect(foldedCaptureNodes(joined.doc)).toHaveLength(0);
    const rebuilt = await expectHonest(joined.doc, parsed, ["code: LISTNUM "]);
    expect(paragraphTokens(rebuilt)).toEqual([["text:Plain.", ...SIMPLE_FIELD_TOKENS, "text:cd"]]);
    expect(
      fieldResultsShown(bodyParagraphs(rebuilt)[0] ?? { type: "paragraph", content: [] }),
    ).toBe("(a)");
  });

  test("the save alone does the same when nothing watched the edit", async () => {
    const { parsed, bare } = await fixture();
    const { node, position } = paragraphNode(bare.doc, LEADING.paraId);
    const unnumbered = bare.apply(
      bare.tr.setNodeMarkup(position, undefined, {
        ...node.attrs,
        ...CLEARED_LIST_RENDERING_ATTRS,
        numPr: null,
      }),
    );
    const joined = unnumbered.apply(
      unnumbered.tr.join(position + paragraphNode(unnumbered.doc, LEADING.paraId).node.nodeSize),
    );

    const rebuilt = fromProseDoc(joined.doc, parsed);

    expectShownIsWritten(rebuilt);
    expect(foldedKinds(rebuilt).at(0)).toEqual([]);
    expect(paragraphTokens(rebuilt).flatMap(codes)).toEqual(ALL_CODES);
    const reopened = await openDocx(await repackDocx(rebuilt, { updateModifiedDate: false }));
    expectShownIsWritten(reopened);
  });
});

describe("moving and copying a paragraph whose marker shows a field", () => {
  const paragraphSlice = (state: EditorState) => {
    const { node, position } = paragraphNode(state.doc, SIMPLE.paraId);
    return {
      position,
      size: node.nodeSize,
      slice: state.doc.slice(position, position + node.nodeSize),
    };
  };

  test("a copy pasted elsewhere shows its field, and the paragraph it came from keeps its own", async () => {
    const { parsed, state } = await fixture([SIMPLE]);
    const { slice } = paragraphSlice(state);
    const end = state.doc.content.size;

    const pasted = state.apply(
      state.tr
        .replaceRange(end, end, unfoldPastedListNumberFields(slice))
        .setNodeAttribute(end, "paraId", FRESH_PARAGRAPH_ID),
    );

    // One hidden field, as before: the copy brought none.
    expect(foldedCaptureNodes(pasted.doc).map(({ node }) => node.attrs["xml"])).toEqual(
      foldedCaptureNodes(state.doc).map(({ node }) => node.attrs["xml"]),
    );
    const rebuilt = await expectHonest(pasted.doc, parsed, ["code: LISTNUM ", "code: LISTNUM "]);
    expect(foldedKinds(rebuilt)).toEqual([["folded:field", "folded:tab"], [], []]);
    expect(paragraphTokens(rebuilt).at(2)).toEqual([...SIMPLE_FIELD_TOKENS, "text:cd"]);
  });

  test("cut and pasted elsewhere, the paragraph still has its field, once, and on the line", async () => {
    const { parsed, state } = await fixture([SIMPLE]);
    const { position, size, slice } = paragraphSlice(state);
    const cut = state.apply(state.tr.delete(position, position + size));
    const end = cut.doc.content.size;

    const pasted = cut.apply(cut.tr.replaceRange(end, end, unfoldPastedListNumberFields(slice)));

    expect(foldedCaptureNodes(pasted.doc)).toHaveLength(0);
    const rebuilt = await expectHonest(pasted.doc, parsed, ["code: LISTNUM "]);
    expect(paragraphTokens(rebuilt)).toEqual([
      ["text:Plain."],
      [...SIMPLE_FIELD_TOKENS, "text:cd"],
    ]);
  });

  test("dragged elsewhere in one step, the paragraph still has its field, once", async () => {
    const { parsed, state } = await fixture([SIMPLE]);
    const { position, size, slice } = paragraphSlice(state);
    const end = state.doc.content.size;

    // A drop deletes the dragged range and inserts the slice in one transaction.
    const tr = state.tr.delete(position, position + size);
    const target = tr.mapping.map(end);
    const dropped = state.apply(
      tr.replaceRange(target, target, unfoldPastedListNumberFields(slice)),
    );

    expect(foldedCaptureNodes(dropped.doc)).toHaveLength(0);
    const rebuilt = await expectHonest(dropped.doc, parsed, ["code: LISTNUM "]);
    expect(paragraphTokens(rebuilt)).toEqual([
      ["text:Plain."],
      [...SIMPLE_FIELD_TOKENS, "text:cd"],
    ]);
  });

  test("a slice that holds no field is pasted as it is", async () => {
    const { state } = await fixture();
    const { node, position } = paragraphNode(state.doc, PLAIN_PARAGRAPH_ID);
    const copied = state.doc.slice(position, position + node.nodeSize);

    expect(unfoldPastedListNumberFields(copied)).toBe(copied);
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
