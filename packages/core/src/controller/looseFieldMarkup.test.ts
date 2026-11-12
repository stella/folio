/**
 * Field markup the parser cannot assemble into a field (an instruction with no
 * field around it, a field whose start sits inside a tracked deletion, a field
 * end a link holds alone, a field the paragraph begins and a later one ends)
 * has no model, only bytes. Both
 * engines keep those bytes through an edit elsewhere in the paragraph.
 */

import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  createHarnessState,
  parseShapeDocument,
  saveHarnessState,
} from "../__tests__/editorHarness";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createEmptyDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import { createCanonicalSession } from "./canonicalSession";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const run = (inner: string) => `<w:r>${inner}</w:r>`;
const text = (value: string) => run(`<w:t xml:space="preserve">${value}</w:t>`);
const char = (type: string) => run(`<w:fldChar w:fldCharType="${type}"/>`);
const instruction = (code: string) =>
  run(`<w:instrText xml:space="preserve">${code}</w:instrText>`);
const REVISION = 'w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';
const EDITED = text("Edit here. ");

/** Paragraph bodies, each after the edited text; some need a second paragraph. */
const SHAPES = {
  "an instruction with no field": [instruction(" \\* MERGEFORMAT ")],
  "a field whose start a tracked deletion holds": [
    `<w:del ${REVISION}>${char("begin")}${run('<w:delInstrText xml:space="preserve"> REF a </w:delInstrText>')}</w:del>`,
    instruction(" \\* MERGEFORMAT "),
    char("separate"),
    text("result"),
    char("end"),
  ],
  "a field end a link holds with no field": [
    `<w:hyperlink w:anchor="a">${text("link")}${char("end")}</w:hyperlink>`,
  ],
  "a field this paragraph begins and the next ends": [
    char("begin"),
    instruction(" TOC \\o "),
  ],
} as const;

const sourceBytes = async (body: readonly string[]) => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body><w:p>${EDITED}${body.join("")}</w:p><w:p>${char("separate")}${text("Contents")}${char("end")}</w:p><w:sectPr/></w:body></w:document>`,
  );
  return zip.generate({ compression: "store" });
};

/** The field markup of the first paragraph, in source order and as written. */
const fieldMarkup = async (bytes: ArrayBuffer | Uint8Array) => {
  const xml =
    (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
    "";
  const paragraph =
    /<w:p[ >].*?<\/w:p>/su.exec(xml)?.[0] ?? panic("No paragraph");
  return [
    ...paragraph.matchAll(
      /<w:fldChar\b[^>]*\/>|<w:(?:del)?instrText\b[^>]*>[^<]*<\/w:(?:del)?instrText>/gu,
    ),
  ].map(([markup]) => markup);
};

/** Where "Edit here" starts in an editor document. */
const editAt = (doc: EditorState["doc"]) => {
  let found: number | undefined;
  doc.descendants((node, position) => {
    if (found === undefined && node.isText && node.text?.startsWith("Edit"))
      found = position;
    return found === undefined;
  });
  return found ?? panic("No edited text");
};

describe("field markup with no assembled field", () => {
  for (const [shape, body] of Object.entries(SHAPES)) {
    test(`${shape} keeps its bytes through an edit in the legacy editor`, async () => {
      const bytes = await sourceBytes(body);
      const base = await parseShapeDocument(bytes);
      const state = createHarnessState(base, "editing");
      const edited = state.apply(state.tr.insertText("Z", editAt(state.doc)));
      const saved = await saveHarnessState(edited, base);
      expect(await fieldMarkup(saved.bytes)).toEqual(await fieldMarkup(bytes));
    });

    test(`${shape} opens and keeps its bytes through an edit in the canonical session`, async () => {
      const bytes = await sourceBytes(body);
      const session = createCanonicalSession(
        await parseShapeDocument(
          (await prepareCanonicalDocxInput(bytes)).unwrap(),
        ),
      ).unwrap();
      let state = EditorState.create({ schema, doc: session.projection.doc });
      const caret = editAt(state.doc);
      state = state.apply(
        state.tr.setSelection(TextSelection.create(state.doc, caret)),
      );
      const commit = session
        .prepareReplace(state, { from: caret, to: caret, text: "Z" })
        .unwrap();
      expect(commit.publish().isOk()).toBe(true);
      const saved = await serializeCanonicalSave({
        snapshot: session.captureSaveSnapshot(),
        featureFlags: { selectiveSave: true },
      });
      expect(await fieldMarkup(saved.buffer)).toEqual(await fieldMarkup(bytes));
    });
  }
});
