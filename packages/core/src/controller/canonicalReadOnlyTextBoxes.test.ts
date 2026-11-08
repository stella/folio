import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";
import {
  type CanonicalCommit,
  type CanonicalSession,
  createCanonicalSession,
} from "./canonicalSession";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="${R}" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
const BOX_TEXT =
  '<w:p w14:paraId="20000001"><w:r><w:t>Inside</w:t></w:r></w:p>';
const drawing = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="4000000" cy="2000000"/><wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${BOX_TEXT}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const HOSTS = {
  // Text beside the box: the paragraph projects, with the box after it.
  inline: `<w:p w14:paraId="10000001"><w:r><w:t>Host</w:t></w:r>${drawing}</w:p>`,
  // Only the box: the box stands where its paragraph was.
  standalone: `<w:p w14:paraId="10000001">${drawing}</w:p>`,
} as const;

const boxedPackage = async (host: string) => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${host}` +
      '<w:p w14:paraId="10000002"><w:r><w:t>Next</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const session = async (
  host: keyof typeof HOSTS,
  mode: "editing" | "suggesting" = "editing",
) => {
  const parsed = await parseDocx(
    (await prepareCanonicalDocxInput(await boxedPackage(HOSTS[host]))).unwrap(),
    { preloadFonts: false, detectVariables: false },
  );
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

const savedXml = async (current: CanonicalSession) => {
  const saved = await serializeCanonicalSave({
    snapshot: current.captureSaveSnapshot(),
    featureFlags: { selectiveSave: true },
  });
  const zip = await ZipArchive.read(saved.buffer);
  return (await zip.entry("word/document.xml")?.text()) ?? "";
};
const boxCount = (xml: string) => xml.split("<w:txbxContent>").length - 1;

describe("text boxes in the canonical session are read-only", () => {
  for (const host of ["inline", "standalone"] as const) {
    test(`a box ${host === "inline" ? "beside text" : "on its own"} opens, and typing beside it saves the box as it was`, async () => {
      const current = await session(host);
      let state = stateFor(current);
      const end = at(state, "Next");
      state = select(state, end);
      state = accept(
        state,
        current
          .prepareReplace(state, { from: end, to: end, text: "!" })
          .unwrap(),
      );
      const xml = await savedXml(current);
      expect(xml).toContain("Next!");
      expect(boxCount(xml)).toBe(1);
      expect(xml).toContain(`<w:txbxContent>${BOX_TEXT}</w:txbxContent>`);
    });

    test(`typing inside a box ${host === "inline" ? "beside text" : "on its own"} refuses and changes nothing`, async () => {
      const current = await session(host);
      const before = current.document;
      let state = stateFor(current);
      const inside = at(state, "Inside", 3);
      state = select(state, inside);
      const typed = current.prepareReplace(state, {
        from: inside,
        to: inside,
        text: "x",
      });
      expect(typed.isErr()).toBe(true);
      expect(current.document).toEqual(before);
    });

    test(`a selection reaching into a box ${host === "inline" ? "beside text" : "on its own"} refuses`, async () => {
      const current = await session(host);
      const before = current.document;
      let state = stateFor(current);
      const from = at(state, "Inside", 3);
      const to = at(state, "Next", 2);
      state = select(state, from, to);
      expect(
        current.prepareReplace(state, { from, to, text: "" }).isErr(),
      ).toBe(true);
      expect(current.document).toEqual(before);
    });
  }

  test("typing in the paragraph that holds a box rewrites it with the box kept", async () => {
    const current = await session("inline");
    let state = stateFor(current);
    const end = at(state, "Host");
    state = select(state, end);
    state = accept(
      state,
      current.prepareReplace(state, { from: end, to: end, text: "!" }).unwrap(),
    );
    const xml = await savedXml(current);
    expect(xml).toContain("Host!");
    expect(xml).toContain(`<w:txbxContent>${BOX_TEXT}</w:txbxContent>`);
    expect(boxCount(xml)).toBe(1);
  });

  for (const mode of ["editing", "suggesting"] as const)
    test(`Backspace after a standalone box in ${mode} mode refuses and changes nothing`, async () => {
      const current = await session("standalone", mode);
      const before = current.document;
      let state = stateFor(current);
      state = select(state, at(state, "Next", 0));
      expect(current.prepareJoin(state, "backward").isErr()).toBe(true);
      expect(current.document).toEqual(before);
    });

  for (const mode of ["editing", "suggesting"] as const)
    test(`a deletion over a whole box in ${mode} mode matches its projection, and undo is exact`, async () => {
      const current = await session("inline", mode);
      const before = current.document;
      let state = stateFor(current);
      const from = at(state, "Host", 2);
      const to = at(state, "Next", 2);
      state = select(state, from, to);
      state = accept(
        state,
        current.prepareReplace(state, { from, to, text: "" }).unwrap(),
      );
      expect(state.doc.eq(current.projection.doc)).toBe(true);
      // Editing removes the selected box with its anchor; a suggestion keeps
      // it, marked deleted, until it is resolved.
      expect(boxCount(await savedXml(current))).toBe(
        mode === "editing" ? 0 : 1,
      );
      state = accept(state, current.prepareUndo(state).unwrap());
      expect(current.document).toEqual(before);
    });
});
