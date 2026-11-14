import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  type CanonicalCommit,
  type CanonicalSession,
  createCanonicalSession,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createEmptyDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CHUNK = '<w:altChunk r:id="rIdChunk"/>';

/** An imported HTML chunk between two paragraphs: markup the editor keeps opaque. */
const chunkedPackage = async () => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>` +
      "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
      CHUNK +
      "<w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>",
  );
  zip.set(
    "word/_rels/document.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdChunk" Type="${R}/aFChunk" Target="afchunk.htm"/>` +
      "</Relationships>",
  );
  zip.set(
    "[Content_Types].xml",
    (await zip.entry("[Content_Types].xml")?.text())?.replace(
      "</Types>",
      '<Override PartName="/word/afchunk.htm" ContentType="text/html"/></Types>',
    ) ?? "",
  );
  zip.set("word/afchunk.htm", "<html><body>Imported clause</body></html>");
  return (await zip.generate({ compression: "store" })).buffer;
};

const session = async (mode: "editing" | "suggesting" = "editing") => {
  const parsed = await parseDocx(
    (await prepareCanonicalDocxInput(await chunkedPackage())).unwrap(),
    { preloadFonts: false, detectVariables: false },
  );
  expect(parsed.package.document.content.map(({ type }) => type)).toEqual([
    "paragraph",
    "preservedBlock",
    "paragraph",
  ]);
  const created = createCanonicalSession(parsed).unwrap();
  if (mode === "suggesting")
    created.setMode({ type: "suggesting", author: "Reviewer" });
  return created;
};

const stateFor = (current: CanonicalSession) =>
  EditorState.create({ schema, doc: current.projection.doc });
const accept = (state: EditorState, commit: CanonicalCommit) => {
  const next = state.apply(commit.transaction);
  expect(commit.publish().isOk()).toBe(true);
  return next;
};
const at = (state: EditorState, text: string, offset = text.length) => {
  let found: number | undefined;
  state.doc.descendants((node, position) => {
    if (found !== undefined) return false;
    if (node.isText && node.text === text) found = position + offset;
    return true;
  });
  return found ?? panic(`No text ${text}`);
};
const select = (state: EditorState, from: number, to = from) =>
  state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));

const savedDocumentXml = async (current: CanonicalSession) => {
  const saved = await serializeCanonicalSave({
    snapshot: current.captureSaveSnapshot(),
    featureFlags: { selectiveSave: true },
  });
  const zip = await ZipArchive.read(saved.buffer);
  return {
    xml: (await zip.entry("word/document.xml")?.text()) ?? "",
    chunk: await zip.entry("word/afchunk.htm")?.text(),
  };
};

describe("opaque blocks in the canonical session", () => {
  test("open, and typing beside one saves it as it was", async () => {
    const current = await session();
    let state = stateFor(current);
    const end = at(state, "Before");
    state = select(state, end);
    state = accept(
      state,
      current.prepareReplace(state, { from: end, to: end, text: "!" }).unwrap(),
    );
    const { xml, chunk } = await savedDocumentXml(current);
    expect(xml.split(CHUNK)).toHaveLength(2);
    expect(xml).toContain("Before!");
    expect(chunk).toBe("<html><body>Imported clause</body></html>");
  });

  for (const mode of ["editing", "suggesting"] as const)
    test(`a deletion across one in ${mode} mode keeps it, and undo is exact`, async () => {
      const current = await session(mode);
      const before = current.document;
      let state = stateFor(current);
      const from = at(state, "Before", 3);
      const to = at(state, "After", 2);
      state = select(state, from, to);
      state = accept(
        state,
        current.prepareReplace(state, { from, to, text: "" }).unwrap(),
      );
      expect(
        current.document.package.document.content.map(({ type }) => type),
      ).toEqual(["paragraph", "preservedBlock", "paragraph"]);
      const { xml } = await savedDocumentXml(current);
      expect(xml.split(CHUNK)).toHaveLength(2);
      state = accept(state, current.prepareUndo(state).unwrap());
      expect(current.document).toEqual(before);
    });
});
