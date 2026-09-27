import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { DOMParser, Slice } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { createDocx } from "../../../docx/rezip";
import { parseDocx } from "../../../docx/parser";
import { completeNumberingForDoc } from "../../listInstanceReferences";
import { schema } from "../../schema";
import { fromProseDoc } from "../../conversion/fromProseDoc";
import { toProseDoc } from "../../conversion/toProseDoc";
import { ParaIdAllocatorExtension } from "./ParaIdAllocatorExtension";
import { flattenPastedHtmlLists, numberPastedHtmlLists } from "./pastedHtmlLists";

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const pasted = (html: string) => {
  const host = document.createElement("div");
  host.innerHTML = flattenPastedHtmlLists(html);
  const parsed = DOMParser.fromSchema(schema).parse(host);
  const state = EditorState.create({
    schema,
    doc: schema.node("doc", null, [schema.node("paragraph", null, [schema.text("Existing")])]),
  });
  const slice = numberPastedHtmlLists(new Slice(parsed.content, 0, 0), { state } as EditorView);
  return schema.node("doc", null, slice.content);
};

describe("HTML list paste", () => {
  test("preserves bullet membership, nested level, and text order", () => {
    const doc = pasted(
      "<ul><li>First bullet</li><li>Second bullet<ul><li>Nested bullet</li></ul></li></ul>",
    );
    expect(doc.childCount).toBe(3);
    expect(
      Array.from({ length: doc.childCount }, (_, index) => doc.child(index).textContent),
    ).toEqual(["First bullet", "Second bullet", "Nested bullet"]);
    const refs = Array.from(
      { length: doc.childCount },
      (_, index) => doc.child(index).attrs["numPr"],
    );
    expect(refs.map((reference) => reference?.ilvl)).toEqual([0, 0, 1]);
    expect(refs[0]?.numId).toBe(refs[1]?.numId);
    expect(doc.child(2).attrs["listIsBullet"]).toBe(true);
    expect(completeNumberingForDoc(undefined, doc)?.nums).toHaveLength(2);
  });

  test("keeps ordered list starts, inline formatting, and surrounding blocks", () => {
    const doc = pasted(
      '<p>Before</p><ol start="4"><li><strong>One</strong></li><li>Two</li></ol><p>After</p>',
    );
    expect(doc.childCount).toBe(4);
    expect(doc.child(0).attrs["numPr"]).toBeNull();
    expect(doc.child(1).attrs["listStartOverride"]).toBe(4);
    expect(doc.child(1).firstChild?.marks.some((mark) => mark.type.name === "bold")).toBe(true);
    expect(doc.child(2).attrs["numPr"]?.numId).toBe(doc.child(1).attrs["numPr"]?.numId);
    expect(doc.child(3).attrs["numPr"]).toBeNull();
  });

  test("defines every pasted list instance in the saved DOCX model", async () => {
    const doc = pasted("<ul><li>Alpha<ul><li>Nested</li></ul></li><li>Omega</li></ul>");
    const saved = fromProseDoc(doc);
    const referenced = new Set<number>();
    doc.forEach((paragraph) => {
      const numId = paragraph.attrs["numPr"]?.numId;
      if (typeof numId === "number") {
        referenced.add(numId);
      }
    });
    expect(new Set(saved.package.numbering?.nums.map(({ numId }) => numId))).toEqual(referenced);
    expect((await createDocx(saved)).byteLength).toBeGreaterThan(0);
  });

  test("saves nested list paste at a caret in an existing document", async () => {
    const bytes = await Bun.file(
      new URL("../../../../../../tests/visual/fixtures/sample.docx", import.meta.url),
    ).arrayBuffer();
    const base = await parseDocx(bytes, { preloadFonts: false, detectVariables: false });
    const plugins = ParaIdAllocatorExtension().onSchemaReady({ schema }).plugins;
    const state = EditorState.create({ schema, doc: toProseDoc(base), plugins });
    const host = document.createElement("div");
    host.innerHTML = flattenPastedHtmlLists(
      "<ul><li>Alpha<ul><li>Nested</li></ul></li><li>Omega</li></ul>",
    );
    const slice = numberPastedHtmlLists(DOMParser.fromSchema(schema).parseSlice(host), {
      state,
    } as EditorView);
    expect([slice.openStart, slice.openEnd]).toEqual([0, 0]);
    const atCaret = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 7)));
    const pastedDoc = atCaret.apply(atCaret.tr.replaceSelection(slice)).doc;
    const listParagraphs: { text: string; level: number }[] = [];
    pastedDoc.forEach((paragraph) => {
      const level = paragraph.attrs["numPr"]?.ilvl;
      if (typeof level === "number") {
        listParagraphs.push({ text: paragraph.textContent, level });
      }
    });
    expect(listParagraphs).toEqual([
      { text: "Alpha", level: 0 },
      { text: "Nested", level: 1 },
      { text: "Omega", level: 0 },
    ]);
    const saved = fromProseDoc(pastedDoc, base);
    expect((await createDocx(saved)).byteLength).toBeGreaterThan(0);
  });
});
