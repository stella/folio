import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createEmptyDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import {
  type CanonicalCommit,
  type CanonicalSession,
  createCanonicalSession,
} from "./canonicalSession";

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";

const run = (inner: string) => `<w:r>${inner}</w:r>`;
const text = (value: string) => run(`<w:t>${value}</w:t>`);
const PAGE_REF = [
  run('<w:fldChar w:fldCharType="begin"/>'),
  run('<w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText>'),
  run('<w:fldChar w:fldCharType="separate"/>'),
  text("2"),
  run('<w:fldChar w:fldCharType="end"/>'),
].join("");
/** A table of contents entry: one link over its text, a tab and the page number. */
const LINK = `<w:hyperlink w:anchor="_Toc1" w:history="1">${text("Summary")}${run("<w:tab/>")}${PAGE_REF}</w:hyperlink>`;

const sourceBytes = async (): Promise<Uint8Array> => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}" xmlns:w14="${W14_NAMESPACE}"><w:body><w:p w14:paraId="10000000">${LINK}</w:p><w:p w14:paraId="20000000">${text("Summary")}</w:p><w:sectPr/></w:body></w:document>`,
  );
  return zip.generate({ compression: "store" });
};

const session = async (mode: "editing" | "suggesting") => {
  const created = createCanonicalSession(
    await parseDocx(
      (await prepareCanonicalDocxInput(await sourceBytes())).unwrap(),
      { preloadFonts: false, detectVariables: false },
    ),
  ).unwrap();
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
/** The first position of the text node `value` in the entry paragraph, plus `offset`. */
const at = (state: EditorState, value: string, offset: number) => {
  let found: number | undefined;
  state.doc.descendants((node, position) => {
    if (found !== undefined) return false;
    if (node.isText && node.text === value) found = position + offset;
    return true;
  });
  return found ?? panic(`No text ${value}`);
};
const select = (state: EditorState, from: number, to = from) =>
  state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));

const savedBytes = async (current: CanonicalSession) =>
  (
    await serializeCanonicalSave({
      snapshot: current.captureSaveSnapshot(),
      featureFlags: { selectiveSave: true },
    })
  ).buffer;
const documentXml = async (bytes: ArrayBuffer | Uint8Array) =>
  (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
  "";
/** Every saved link, with run properties left out. */
const savedLinks = async (bytes: ArrayBuffer | Uint8Array) =>
  [
    ...(await documentXml(bytes)).matchAll(
      /<w:hyperlink\b[^>]*>.*?<\/w:hyperlink>/gsu,
    ),
  ].map(([link]) => link.replaceAll(/<w:rPr>.*?<\/w:rPr>/gu, ""));
/** The saved link holding the page number. */
const fieldLink = async (bytes: ArrayBuffer | Uint8Array) =>
  (await savedLinks(bytes)).find((link) => link.includes("PAGEREF"));
/**
 * The link after the caret, as typing in the entry leaves it: the typed text
 * lands between two links, and the second keeps the tab and the page number.
 */
const TAIL_LINK = `<w:hyperlink w:anchor="_Toc1" w:history="1">${text("mary")}${run("<w:tab/>")}${PAGE_REF}</w:hyperlink>`;

describe("a field inside a hyperlink in the canonical session", () => {
  test("opens, and typing in the entry keeps the page number field inside its link", async () => {
    const current = await session("editing");
    let state = stateFor(current);
    const caret = at(state, "Summary", 3);
    state = select(state, caret);
    state = accept(
      state,
      current
        .prepareReplace(state, { from: caret, to: caret, text: "X" })
        .unwrap(),
    );
    expect(await fieldLink(await savedBytes(current))).toBe(TAIL_LINK);
  });

  test("a suggestion in the entry keeps the page number field inside its link, and rejects to the source link", async () => {
    const current = await session("suggesting");
    let state = stateFor(current);
    const caret = at(state, "Summary", 3);
    state = select(state, caret);
    state = accept(
      state,
      current
        .prepareReplace(state, { from: caret, to: caret, text: "X" })
        .unwrap(),
    );
    const saved = await savedBytes(current);
    for (const resolution of ["accept", "reject"] as const) {
      const reviewer = await FolioDocxReviewer.fromBuffer(saved);
      if (resolution === "accept") reviewer.acceptAll();
      else reviewer.rejectAll();
      const resolved = await reviewer.toBuffer();
      // Rejecting rejoins the link the suggestion cut, field and all.
      expect(await fieldLink(resolved)).toBe(
        resolution === "accept" ? TAIL_LINK : LINK,
      );
    }
  });

  test("deleting from the entry text through the page number removes the whole field", async () => {
    const current = await session("editing");
    let state = stateFor(current);
    const from = at(state, "Summary", 3);
    let paragraphEnd = 0;
    state.doc.forEach((node, position) => {
      if (paragraphEnd === 0) paragraphEnd = position + node.nodeSize - 1;
    });
    state = select(state, from, paragraphEnd);
    state = accept(
      state,
      current
        .prepareReplace(state, { from, to: paragraphEnd, text: "" })
        .unwrap(),
    );
    const xml = await documentXml(await savedBytes(current));
    expect(xml).not.toContain("fldChar");
    expect(xml).not.toContain("PAGEREF");
    expect(await savedLinks(await savedBytes(current))).toEqual([
      `<w:hyperlink w:anchor="_Toc1" w:history="1">${text("Sum")}</w:hyperlink>`,
    ]);
  });
});
