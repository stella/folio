/**
 * A comment whose covered text is all deleted retires, also when bookmarks or
 * other range boundaries sat inside its range: they mark positions, not
 * content. Both engines, after a save.
 */
import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const COMMENTS_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const fixture = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId9" Type="${COMMENTS_TYPE}" Target="comments.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      `<w:p w14:paraId="1A2B3C4D"><w:commentRangeStart w:id="0"/><w:bookmarkStart w:id="5" w:name="_GoBack"/>` +
      `${run("This")}<w:bookmarkEnd w:id="5"/><w:commentRangeEnd w:id="0"/>` +
      `<w:r><w:commentReference w:id="0"/></w:r>${run(" stays.")}</w:p>` +
      `<w:p w14:paraId="2B3C4D5E">${run("Tail.")}</w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  zip.set(
    "word/comments.xml",
    `${HEAD}<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="Reviewer" w:date="2024-01-01T00:00:00Z" w:initials="R">` +
      `<w:p>${run("Check this.")}</w:p></w:comment></w:comments>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const DELETE_COMMENTED_TEXT = {
  id: "drop",
  type: "replaceInBlock",
  blockId: "1A2B3C4D",
  find: "This",
  replace: "",
} as const;

test("the legacy engine drops the comment from the save", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await fixture());
  expect(reviewer.getComments()).toHaveLength(1);
  const result = reviewer.applyOperations([DELETE_COMMENTED_TEXT], {
    undefinedReferences: "refuse",
    mode: "direct",
  });
  expect(result.skipped).toEqual([]);
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
  );
  expect(reopened.getComments()).toEqual([]);
  expect(reopened.getContent().map(({ text }) => text)).toContain(" stays.");
});

test("the canonical session drops the comment from the save", async () => {
  const prepared = (await prepareCanonicalDocxInput(await fixture())).unwrap();
  const document = await parseDocx(prepared, {
    preloadFonts: false,
    detectVariables: false,
  });
  const session = createCanonicalSession(
    document,
    document.package.styles,
  ).unwrap();
  let state = EditorState.create({ doc: session.projection.doc });
  const publish = (commit: CanonicalCommit) => {
    const published = publishCanonicalProjection({ state, commit, session });
    if (published.isErr()) return false;
    state = published.value.state;
    return true;
  };
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish,
  });
  const result = executor.apply({
    snapshot: createFolioAIEditSnapshot(state.doc),
    batch: { version: 1, mode: "direct", operations: [DELETE_COMMENTED_TEXT] },
    undefinedReferences: "refuse",
  });
  expect(result.applied).toHaveLength(1);
  const saved = await serializeCanonicalSave({
    snapshot: session.captureSaveSnapshot(),
  });
  const reopened = await FolioDocxReviewer.fromBuffer(saved.buffer);
  expect(reopened.getComments()).toEqual([]);
  expect(reopened.getContent().map(({ text }) => text)).toContain(" stays.");
});

test("a source comment whose range holds only a bookmark stays a point comment", async () => {
  const zip = await ZipArchive.read(await fixture());
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  zip.set("word/document.xml", xml.replace(run("This"), ""));
  const reviewer = await FolioDocxReviewer.fromBuffer(
    (await zip.generate({ compression: "store" })).buffer,
  );
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
  );
  expect(reopened.getComments()).toHaveLength(1);
});
