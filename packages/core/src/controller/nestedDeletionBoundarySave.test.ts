import { expect, setDefaultTimeout, test } from "bun:test";
import { type Command, EditorState } from "prosemirror-state";

import { validateDocxPackage } from "@stll/docx-core";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import { propertyTestTimeout } from "../../../../test/property-testing";
import { createCanonicalSession } from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { createDocx, repackDocx } from "../docx/rezip";
import { createSimpleDocument } from "../docx/serializer/documentSerializer";
import {
  acceptAllChanges,
  rejectAIEditRevision,
  rejectAllChanges,
} from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import type { Document } from "../types/document";
import { FOLIO_DOCX_SERIALIZATION_MODE } from "../types/docxSerialization";

setDefaultTimeout(propertyTestTimeout(30_000));

// Rejecting an insertion that holds another reviewer's pending deletion of a
// bookmark leaves that deletion pending, holding only the bookmark. Such a
// deletion is ordinary tracked content: it saves, reopens and resolves as any
// other.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const PARAGRAPH = "1A2B3C4D";

const source = async (body: string): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(
    await createDocx(createSimpleDocument([{ text: "seed" }])),
  );
  zip.set(
    "word/document.xml",
    `<?xml version="1.0"?><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body><w:p w14:paraId="${PARAGRAPH}"><w:r><w:t>a</w:t></w:r>${body}<w:r><w:t>b</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const nested = () =>
  source(
    '<w:ins w:id="41" w:author="A" w:date="2026-03-04T05:00:01Z"><w:del w:id="42" w:author="B" w:date="2026-03-04T05:00:02Z"><w:bookmarkStart w:id="5" w:name="_Ref5"/><w:bookmarkEnd w:id="5"/></w:del></w:ins>',
  );

const open = (buffer: ArrayBuffer) =>
  parseDocx(buffer, { preloadFonts: false });

const paragraphContent = (document: Document) => {
  const first = document.package.document.content[0];
  return first?.type === "paragraph" ? first.content : [];
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await ZipArchive.read(buffer)).entry("word/document.xml")?.text()) ??
  "";

const bookmark = [
  { type: "bookmarkStart", id: 5, name: "_Ref5" },
  { type: "bookmarkEnd", id: 5 },
];

/** The paragraph's content once `command` ran in the editor and the result was saved. */
const savedAfter = async (document: Document, command: Command) => {
  const state = EditorState.create({ schema, doc: toProseDoc(document) });
  let resolved = state.doc;
  expect(command(state, (tr) => (resolved = tr.doc))).toBe(true);
  const saved = await repackDocx(
    fromProseDoc(resolved, document, { stylesheetSource: { type: "package" } }),
    { updateModifiedDate: false },
  );
  expect(await validateDocxPackage(new Uint8Array(saved))).toEqual({
    valid: true,
  });
  return saved;
};

test("the editor's rejection saves the deletion holding only the bookmark, and it reopens", async () => {
  const saved = await savedAfter(
    await open(await nested()),
    rejectAIEditRevision(41),
  );
  const xml = await documentXml(saved);
  expect(xml).not.toContain("<w:ins");
  expect(xml).toMatch(
    /<w:del [^>]*w:id="42"[^>]*><w:bookmarkStart [^>]*w:id="5"[^>]*\/><w:bookmarkEnd [^>]*w:id="5"[^>]*\/><\/w:del>/u,
  );
  const reopened = await open(saved);
  expect(paragraphContent(reopened)).toMatchObject([
    { type: "run" },
    { type: "deletion", info: { id: 42, author: "B" }, content: bookmark },
    { type: "run" },
  ]);
  // Accepting it removes the bookmark and leaves no wrapper; rejecting keeps it.
  const accepted = await open(await savedAfter(reopened, acceptAllChanges()));
  expect(paragraphContent(accepted).map(({ type }) => type)).toEqual(
    paragraphContent(accepted).map(() => "run"),
  );
  expect(JSON.stringify(paragraphContent(accepted))).not.toContain("bookmark");
  const rejected = await open(await savedAfter(reopened, rejectAllChanges()));
  expect(
    paragraphContent(rejected).filter(({ type }) => type !== "run"),
  ).toEqual(bookmark);
});

test("the canonical save writes the same deletion and passes its checks", async () => {
  const session = createCanonicalSession(await open(await nested())).unwrap();
  const state = EditorState.create({ schema, doc: session.projection.doc });
  const commit = session
    .prepareOperations(state, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [41],
        decision: REVISION_DECISIONS.REJECT,
      },
    ])
    .unwrap();
  commit.publish().unwrap();
  const snapshot = session.captureSaveSnapshot();
  for (const mode of Object.values(FOLIO_DOCX_SERIALIZATION_MODE)) {
    const saved = await serializeCanonicalSave({ snapshot, options: { mode } });
    expect(await validateDocxPackage(new Uint8Array(saved.buffer))).toEqual({
      valid: true,
    });
    const reopened = await open(saved.buffer);
    expect(paragraphContent(reopened)).toMatchObject([
      { type: "run" },
      { type: "deletion", info: { id: 42, author: "B" }, content: bookmark },
      { type: "run" },
    ]);
  }
});
