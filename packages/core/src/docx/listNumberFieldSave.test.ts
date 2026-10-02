/**
 * A numbered paragraph's inline `LISTNUM` fields are drawn as part of its list
 * marker. They stay in the paragraph's content as the markup they were read
 * from, showing nothing, so a save owes the file every one of them as it was:
 * the instruction, the field characters, the cached result with its
 * formatting, the tab after it, and where they sat.
 */

import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { removeFoldedListNumberFields } from "../prosemirror/extensions/features/pasteCleanup";
import { CLEARED_LIST_RENDERING_ATTRS } from "../prosemirror/listMarker";
import { paragraphNumberingAttr } from "../prosemirror/numberingAttr";
import type { Document } from "../types/document";
import {
  bodyParagraphs,
  contentShapes,
  documentXmlOf,
  expectedTokens,
  foldedCaptureNodes,
  inlineTokens,
  layoutMarkers,
  listNumberFieldDocx,
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
      formatting: "symbol",
      before: "first; ",
      gap: ["comment"],
      tab: false,
    },
  ],
  body: " second",
};

const SPECS = [LEADING, INLINE];

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
  expect(paragraphMarkupOf(xml, INLINE.paraId)).toMatch(
    /w:ascii="Symbol"(?:(?!<\/w:r>).)*<w:t[^>]*>\(i\)<\/w:t>/u,
  );

  const reopened = await openDocx(saved);
  expect(foldedKinds(reopened)).toEqual(foldedKinds(original));
  expect(modelMarkers(reopened)).toEqual(modelMarkers(original));
  expect(layoutMarkers(reopened)).toEqual(layoutMarkers(original));
};

const fixture = async (
  specs: readonly ParagraphSpec[] = SPECS,
): Promise<{ buffer: ArrayBuffer; parsed: Document; state: EditorState }> => {
  const buffer = await listNumberFieldDocx(specs);
  const parsed = await openDocx(buffer);
  return { buffer, parsed, state: EditorState.create({ doc: toProseDoc(parsed) }) };
};

describe("the reader's fold of LISTNUM fields", () => {
  test("shows each field in the marker and keeps its markup where it stood", async () => {
    const { parsed } = await fixture();
    const [leading, inline] = bodyParagraphs(parsed);
    if (!leading || !inline) {
      throw new Error("The fixture opens with two numbered paragraphs");
    }

    expect(leading.listRendering?.marker.endsWith("\t(a)")).toBe(true);
    expect(inline.listRendering?.marker.endsWith("\t(b) (i)")).toBe(true);
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
    expect(inlineShape?.slice(0, 6)).toEqual([
      "folded:field",
      "folded:tab",
      "run",
      "folded:field",
      "commentRangeStart",
      "run",
    ]);
    // A capture shows nothing: the marker is where the field is read.
    for (const item of [...leading.content, ...inline.content]) {
      if (item.type === "preservedInline") {
        expect(item.text).toBe("");
      }
    }
    expect(paragraphTokens(parsed).slice(0, 2).map(codes).flat()).toEqual(ALL_CODES);
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
      [LEADING.paraId]:
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
        `<w:r><w:instrText xml:space="preserve"> LISTNUM </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="end"><w:numberingChange w:id="9" w:author="A" w:date="2024-01-01T00:00:00Z" w:original="(a)"/></w:fldChar></w:r>` +
        `<w:r><w:tab/></w:r><w:r><w:t>Body text</w:t></w:r>`,
    });
    const parsed = await openDocx(buffer);
    expect(bodyParagraphs(parsed).at(0)?.listRendering?.marker.endsWith("\t(a)")).toBe(true);
    const edited = typeInto(
      EditorState.create({ doc: toProseDoc(parsed) }),
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

const ANCHORED: ParagraphSpec = {
  paraId: "20000001",
  marker: "decimal",
  fields: [
    {
      instruction: " LISTNUM ",
      result: "(a)",
      formatting: "plain",
      before: "ab ",
      gap: [],
      tab: true,
    },
  ],
  body: "cd",
};

type Edit = (state: EditorState) => EditorState;

const insertAt =
  (position: (state: EditorState) => number, text: string): Edit =>
  (state) =>
    state.apply(state.tr.insertText(text, position(state)));

const deleteText =
  (text: string): Edit =>
  (state) => {
    const from = positionOfText(state.doc, ANCHORED.paraId, text);
    return state.apply(state.tr.delete(from, from + text.length));
  };

const capture = (state: EditorState, kind: "field" | "tab"): number => {
  const found = foldedCaptureNodes(state.doc).find(
    ({ node }) => node.attrs["foldedListNumber"] === kind,
  );
  if (!found) {
    throw new Error(`The paragraph holds no ${kind} capture`);
  }
  return found.position;
};

const startOf =
  (text: string) =>
  (state: EditorState): number =>
    positionOfText(state.doc, ANCHORED.paraId, text);

const endOf =
  (text: string) =>
  (state: EditorState): number =>
    positionOfText(state.doc, ANCHORED.paraId, text) + text.length;

const ANCHOR_CASES: { name: string; edits: Edit[]; before: string; body: string }[] = [
  {
    name: "typed ahead of the field",
    edits: [insertAt(startOf("ab "), "X")],
    before: "Xab ",
    body: "cd",
  },
  { name: "deleted ahead of the field", edits: [deleteText("b")], before: "a ", body: "cd" },
  {
    name: "typed right before the field",
    edits: [insertAt((state) => capture(state, "field"), "Y")],
    before: "ab Y",
    body: "cd",
  },
  {
    name: "deleted right before the field",
    edits: [deleteText(" ")],
    before: "ab",
    body: "cd",
  },
  {
    name: "typed right after the tab",
    edits: [insertAt((state) => capture(state, "tab") + 1, "Z")],
    before: "ab ",
    body: "Zcd",
  },
  { name: "deleted right after the tab", edits: [deleteText("c")], before: "ab ", body: "d" },
  {
    name: "typed after the field",
    edits: [insertAt(endOf("cd"), "W")],
    before: "ab ",
    body: "cdW",
  },
  {
    name: "typed and deleted on every side at once",
    edits: [
      insertAt(startOf("ab "), "X"),
      deleteText("b"),
      insertAt((state) => capture(state, "field"), "Y"),
      insertAt((state) => capture(state, "tab") + 1, "Z"),
      deleteText("c"),
      insertAt(endOf("d"), "W"),
    ],
    before: "Xa Y",
    body: "ZdW",
  },
];

describe("a LISTNUM field in the middle of a paragraph stays between its neighbours", () => {
  for (const { name, edits, before, body } of ANCHOR_CASES) {
    test(`text ${name}`, async () => {
      const { buffer, parsed, state } = await fixture([ANCHORED]);
      let edited = state;
      for (const edit of edits) {
        edited = edit(edited);
      }

      const saved = await saveDocx(fromProseDoc(edited.doc, parsed), buffer, [ANCHORED.paraId]);

      const markup = paragraphMarkupOf(await documentXmlOf(saved), ANCHORED.paraId);
      expect(inlineTokens(markup)).toEqual([
        `text:${before}`,
        "fldChar:begin",
        "code: LISTNUM ",
        "fldChar:separate",
        "text:(a)",
        "fldChar:end",
        "tab",
        `text:${body}`,
      ]);
      expect(modelMarkers(await openDocx(saved))).toEqual(modelMarkers(parsed));
    });
  }
});

describe("editing a paragraph whose list marker holds LISTNUM fields", () => {
  test("a split leaves the fields in the half they stood in", async () => {
    const { parsed, state } = await fixture();
    const middle = positionOfText(state.doc, LEADING.paraId, " text");

    const rebuilt = fromProseDoc(state.apply(state.tr.split(middle)).doc, parsed);

    const [first, second] = paragraphTokens(rebuilt);
    expect(first).toEqual([...LEADING_FIELD_TOKENS, "text:Body"]);
    expect(second).toEqual(["text: text"]);
    expect(paragraphTokens(rebuilt).map(codes).flat()).toEqual(ALL_CODES);
  });

  test("a split ahead of the fields leaves them in the second half, and only there", async () => {
    const { parsed, state } = await fixture();
    const { position } = paragraphNode(state.doc, LEADING.paraId);

    const rebuilt = fromProseDoc(state.apply(state.tr.split(position + 1)).doc, parsed);

    const [first, second] = paragraphTokens(rebuilt);
    expect(first).toEqual([]);
    expect(second).toEqual([...LEADING_FIELD_TOKENS, "text:Body text"]);
    expect(paragraphTokens(rebuilt).map(codes).flat()).toEqual(ALL_CODES);
  });

  test("a join keeps the fields of both paragraphs, each where it stood", async () => {
    const { parsed, state } = await fixture();
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);

    const rebuilt = fromProseDoc(state.apply(state.tr.join(position + node.nodeSize)).doc, parsed);

    expect(bodyParagraphs(rebuilt)).toHaveLength(2);
    const tokens = paragraphTokens(rebuilt).at(0) ?? [];
    expect(codes(tokens)).toEqual(ALL_CODES);
    expect(tokens.slice(0, LEADING_FIELD_TOKENS.length + 2)).toEqual([
      ...LEADING_FIELD_TOKENS,
      "text:Body text",
      "fldChar:begin",
    ]);
    expect(tokens).toContain("text:first; ");
  });

  test("a copy pasted elsewhere takes no field, and the paragraph it came from keeps its own", async () => {
    const { state } = await fixture();
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);
    const copied = state.doc.slice(position, position + node.nodeSize);
    expect(foldedCaptureNodes(copied.content.firstChild ?? node)).toHaveLength(2);

    const pasted = state.apply(state.tr.replaceRange(0, 0, removeFoldedListNumberFields(copied)));

    const copy = pasted.doc.firstChild;
    if (!copy) {
      throw new Error("The paste put a paragraph at the start of the document");
    }
    expect(copy.textContent).toBe(LEADING.body);
    expect(foldedCaptureNodes(copy)).toHaveLength(0);
    expect(foldedCaptureNodes(pasted.doc).map(({ node: found }) => found.attrs["xml"])).toEqual(
      foldedCaptureNodes(state.doc).map(({ node: found }) => found.attrs["xml"]),
    );
  });

  test("a slice that holds no field is pasted as it is", async () => {
    const { state } = await fixture();
    const { node, position } = paragraphNode(state.doc, PLAIN_PARAGRAPH_ID);
    const copied = state.doc.slice(position, position + node.nodeSize);

    expect(removeFoldedListNumberFields(copied)).toBe(copied);
  });

  test("a paragraph moved to another level of its list keeps the fields", async () => {
    const { parsed, state } = await fixture();
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);
    const moved = state.apply(
      state.tr.setNodeMarkup(position, undefined, {
        ...node.attrs,
        numPr: paragraphNumberingAttr(paragraphNumberingReference({ numId: 1, ilvl: 2 })),
      }),
    );

    const tokens = paragraphTokens(fromProseDoc(moved.doc, parsed));

    expect(tokens.at(0)).toEqual([...LEADING_FIELD_TOKENS, "text:Body text"]);
    expect(tokens.map(codes).flat()).toEqual(ALL_CODES);
  });

  test("a paragraph taken out of its list keeps the fields", async () => {
    const { parsed, state } = await fixture();
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);
    const unnumbered = state.apply(
      state.tr.setNodeMarkup(position, undefined, {
        ...node.attrs,
        ...CLEARED_LIST_RENDERING_ATTRS,
        numPr: null,
      }),
    );

    const tokens = paragraphTokens(fromProseDoc(unnumbered.doc, parsed));

    expect(tokens.at(0)).toEqual([...LEADING_FIELD_TOKENS, "text:Body text"]);
  });

  test("a field deleted with the content around it is not written", async () => {
    const { parsed, state } = await fixture();
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);
    const emptied = state.apply(state.tr.delete(position + 1, position + node.nodeSize - 1));

    const tokens = paragraphTokens(fromProseDoc(emptied.doc, parsed));

    expect(tokens.at(0)).toEqual([]);
    expect(tokens.map(codes).flat()).toEqual(ALL_CODES.slice(1));
  });
});
