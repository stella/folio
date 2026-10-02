/**
 * A numbered paragraph's inline `LISTNUM` fields are drawn as part of its list
 * marker, so the paragraph's content does not hold them. A save still owes the
 * file every one of them: the instruction, the field characters, the cached
 * result with its formatting, the tab after it, and where they sat.
 */

import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { CLEARED_LIST_RENDERING_ATTRS } from "../prosemirror/listMarker";
import { paragraphNumberingAttr } from "../prosemirror/numberingAttr";
import type { Document } from "../types/document";
import {
  bodyParagraphs,
  documentXmlOf,
  expectedTokens,
  foldedSource,
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

const codes = (tokens: readonly string[]): string[] =>
  tokens.filter((token) => token.startsWith("code:"));

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
  expect(foldedSource(reopened)).toEqual(foldedSource(original));
  expect(modelMarkers(reopened)).toEqual(modelMarkers(original));
  expect(layoutMarkers(reopened)).toEqual(layoutMarkers(original));
};

const fixture = async (): Promise<{ buffer: ArrayBuffer; parsed: Document }> => {
  const buffer = await listNumberFieldDocx(SPECS);
  return { buffer, parsed: await openDocx(buffer) };
};

describe("the reader's fold of LISTNUM fields", () => {
  test("shows each field in the marker and keeps what it took out of the content", async () => {
    const { parsed } = await fixture();
    const [leading, inline] = bodyParagraphs(parsed);
    if (!leading || !inline) {
      throw new Error("The fixture opens with two numbered paragraphs");
    }

    expect(leading.listRendering?.marker.endsWith("\t(a)")).toBe(true);
    expect(inline.listRendering?.marker.endsWith("\t(b) (i)")).toBe(true);
    expect(JSON.stringify([leading.content, inline.content])).not.toContain("LISTNUM");
    expect(expectedTokens(LEADING).slice(0, LEADING_FIELD_TOKENS.length)).toEqual(
      LEADING_FIELD_TOKENS,
    );

    expect(leading.foldedListNumberFields).toMatchObject({
      numId: 1,
      level: 1,
      fields: [
        {
          field: { instruction: " LISTNUM  LegalDefault \\l 3 ", fieldType: "LISTNUM" },
          tab: { type: "run", content: [{ type: "tab" }] },
          offset: 0,
          markersBefore: 0,
          markersBeforeTab: 2,
        },
      ],
    });
    expect(leading.foldedListNumberFields?.fields).toHaveLength(1);
    expect(inline.foldedListNumberFields?.fields.map(({ offset }) => offset)).toEqual([0, 7]);
    expect(inline.foldedListNumberFields?.fields.map(({ tab }) => tab !== undefined)).toEqual([
      true,
      false,
    ]);
  });
});

describe("saving a paragraph whose list marker holds LISTNUM fields", () => {
  test("an untouched paragraph keeps its fields when another one is edited", async () => {
    const { buffer, parsed } = await fixture();
    const state = typeInto(
      EditorState.create({ doc: toProseDoc(parsed) }),
      PLAIN_PARAGRAPH_ID,
      "Plain.",
      "!",
    );

    const saved = await saveDocx(fromProseDoc(state.doc, parsed), buffer, [PLAIN_PARAGRAPH_ID]);

    await expectSavedFields(saved, parsed);
    expect(await documentXmlOf(saved)).toContain("P!lain.");
  });

  test("an edit elsewhere in the paragraph keeps its fields", async () => {
    const { buffer, parsed } = await fixture();
    let state = EditorState.create({ doc: toProseDoc(parsed) });
    state = typeInto(state, LEADING.paraId, LEADING.body, "!");
    state = typeInto(state, INLINE.paraId, INLINE.body, "!");

    const saved = await saveDocx(fromProseDoc(state.doc, parsed), buffer, [
      LEADING.paraId,
      INLINE.paraId,
    ]);

    await expectSavedFields(saved, parsed, {
      [LEADING.paraId]: typedInto(LEADING.body, "!"),
      [INLINE.paraId]: typedInto(INLINE.body, "!"),
    });
  });

  test("a full rewrite keeps the fields of every paragraph", async () => {
    const { parsed } = await fixture();

    const saved = await repackDocx(fromProseDoc(toProseDoc(parsed), parsed), {
      updateModifiedDate: false,
    });

    await expectSavedFields(saved, parsed);
  });

  test("a second save writes what the first one did", async () => {
    const { parsed } = await fixture();
    const once = await repackDocx(fromProseDoc(toProseDoc(parsed), parsed), {
      updateModifiedDate: false,
    });
    const reopened = await openDocx(once);

    const twice = await repackDocx(fromProseDoc(toProseDoc(reopened), reopened), {
      updateModifiedDate: false,
    });

    expect(await documentXmlOf(twice)).toBe(await documentXmlOf(once));
  });

  test("a marker with no recorded field writes no field", async () => {
    const { parsed } = await fixture();
    for (const paragraph of bodyParagraphs(parsed)) {
      delete paragraph.foldedListNumberFields;
    }

    const rebuilt = fromProseDoc(toProseDoc(parsed), parsed);

    expect(modelMarkers(rebuilt).at(0)?.endsWith("\t(a)")).toBe(true);
    expect(foldedSource(rebuilt)).toEqual([null, null, null]);
    expect(bodyParagraphs(rebuilt).map(serializeParagraph).join("")).not.toContain("fldChar");
  });
});

describe("editing a paragraph whose list marker holds LISTNUM fields", () => {
  test("a split leaves the fields with the first half", async () => {
    const { parsed } = await fixture();
    const source = bodyParagraphs(parsed).at(0)?.foldedListNumberFields;
    const state = EditorState.create({ doc: toProseDoc(parsed) });
    const middle = positionOfText(state.doc, LEADING.paraId, " text");

    const rebuilt = fromProseDoc(state.apply(state.tr.split(middle)).doc, parsed);

    const [first, second] = bodyParagraphs(rebuilt);
    if (!first || !second || source === undefined) {
      throw new Error("The split leaves two halves of a paragraph that folded a field");
    }
    expect(first.foldedListNumberFields).toBe(source);
    expect(second.foldedListNumberFields).toBeUndefined();
    expect(inlineTokens(serializeParagraph(first))).toEqual([...LEADING_FIELD_TOKENS, "text:Body"]);
    expect(inlineTokens(serializeParagraph(second))).toEqual(["text: text"]);
  });

  test("a join keeps the fields of the paragraph the joined one starts with", async () => {
    const { parsed } = await fixture();
    const source = bodyParagraphs(parsed).at(0)?.foldedListNumberFields;
    const state = EditorState.create({ doc: toProseDoc(parsed) });
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);

    const rebuilt = fromProseDoc(state.apply(state.tr.join(position + node.nodeSize)).doc, parsed);

    const joined = bodyParagraphs(rebuilt).at(0);
    if (!joined || source === undefined) {
      throw new Error("The join leaves a paragraph that folded a field");
    }
    expect(bodyParagraphs(rebuilt)).toHaveLength(2);
    expect(joined.foldedListNumberFields).toBe(source);
    const tokens = inlineTokens(serializeParagraph(joined));
    // The second paragraph's fields stood at the start of a paragraph that is
    // gone, so only the first paragraph's are written, and where they stood.
    expect(tokens.slice(0, LEADING_FIELD_TOKENS.length)).toEqual(LEADING_FIELD_TOKENS);
    expect(codes(tokens)).toEqual(["code: LISTNUM  LegalDefault \\l 3 "]);
    expect(tokens).toContain("text:Body textfirst; ");
  });

  test("a paragraph moved to another level of its list drops the fields", async () => {
    const { parsed } = await fixture();
    const state = EditorState.create({ doc: toProseDoc(parsed) });
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);
    const moved = state.apply(
      state.tr.setNodeMarkup(position, undefined, {
        ...node.attrs,
        numPr: paragraphNumberingAttr(paragraphNumberingReference({ numId: 1, ilvl: 2 })),
      }),
    );

    const [first, second] = bodyParagraphs(fromProseDoc(moved.doc, parsed));

    if (!first || !second) {
      throw new Error("The document keeps its two numbered paragraphs");
    }
    expect(first.foldedListNumberFields).toBeUndefined();
    expect(serializeParagraph(first)).not.toContain("fldChar");
    // The paragraph beside it kept its level, and its fields.
    expect(codes(inlineTokens(serializeParagraph(second)))).toEqual([
      "code: LISTNUM ",
      "code:LISTNUM \\l 4",
    ]);
  });

  test("a paragraph taken out of its list drops the fields", async () => {
    const { parsed } = await fixture();
    const state = EditorState.create({ doc: toProseDoc(parsed) });
    const { node, position } = paragraphNode(state.doc, LEADING.paraId);
    const unnumbered = state.apply(
      state.tr.setNodeMarkup(position, undefined, {
        ...node.attrs,
        ...CLEARED_LIST_RENDERING_ATTRS,
        numPr: null,
      }),
    );

    const first = bodyParagraphs(fromProseDoc(unnumbered.doc, parsed)).at(0);

    if (!first) {
      throw new Error("The document keeps its first paragraph");
    }
    expect(first.foldedListNumberFields).toBeUndefined();
    expect(serializeParagraph(first)).not.toContain("fldChar");
    expect(inlineTokens(serializeParagraph(first))).toEqual([
      "bookmarkStart",
      "bookmarkEnd",
      "text:Body text",
    ]);
  });
});
