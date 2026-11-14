/**
 * A table row may hold no cells at all (`CT_Row` takes zero or more cell
 * content items). The canonical session opens a document with one, edits a
 * cell beside it, and saves the row as the file states it: no cell and no
 * paragraph it never had, as the default engine saves it.
 */

import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { CanonicalPublicOperations } from "../canonical/canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const cell = (text: string) =>
  `<w:tc><w:tcPr/><w:p w14:paraId="1000000${text}"><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

/** Rows with no cells, and the omitted positions beside cells. */
const ROWS = {
  "nothing at all": "<w:tr/>",
  "only properties": '<w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr></w:tr>',
  "only a w:gridAfter":
    '<w:tr><w:trPr><w:gridAfter w:val="2"/></w:trPr></w:tr>',
  "only a w:gridBefore":
    '<w:tr><w:trPr><w:gridBefore w:val="2"/></w:trPr></w:tr>',
  "a w:gridBefore and one cell": `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>${cell("5")}</w:tr>`,
  "one cell and a w:gridAfter": `<w:tr><w:trPr><w:gridAfter w:val="1"/></w:trPr>${cell("6")}</w:tr>`,
  "one cell stopping short": `<w:tr>${cell("7")}</w:tr>`,
} as const;

const docx = async (row: string): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr>${cell("1")}${cell("2")}</w:tr>${row}<w:tr>${cell("3")}${cell("4")}</w:tr></w:tbl><w:p w14:paraId="10000009"><w:r><w:t>Tail</w:t></w:r></w:p></w:body></w:document>`,
  );
  return (await zip.generate()).slice().buffer;
};

/** Each saved row's cell count and stated omissions, in order. */
const savedRows = async (buffer: ArrayBuffer) => {
  const xml =
    (await (
      await ZipArchive.read(buffer)
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return [
    ...xml.matchAll(/<w:tr(?:\s[^>]*)?(?:\/>|>([\s\S]*?)<\/w:tr>)/gu),
  ].map(([, inner = ""]) => ({
    cells: (inner.match(/<w:tc[ >]/gu) ?? []).length,
    gridBefore: /<w:gridBefore w:val="(\d+)"/u.exec(inner)?.[1],
    gridAfter: /<w:gridAfter w:val="(\d+)"/u.exec(inner)?.[1],
    text: [...inner.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/gu)]
      .map(([, text]) => text)
      .join(""),
  }));
};

const EDIT = {
  id: "edit",
  type: "replaceInBlock",
  blockId: "10000001",
  find: "1",
  replace: "9",
} as const;

type Mode = "direct" | "tracked-changes";

const viaCanonicalEngine = async (source: ArrayBuffer, mode: Mode) => {
  const document = await parseDocx(source.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  const session = createCanonicalSession(
    document,
    document.package.styles,
  ).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish: (commit: CanonicalCommit) => {
      const published = publishCanonicalProjection({ state, commit, session });
      if (published.isErr()) return false;
      state = published.value.state;
      return true;
    },
  });
  const result = executor.apply({
    snapshot: createFolioAIEditSnapshot(state.doc),
    batch: { version: 1, mode, operations: [EDIT] },
  });
  return {
    applied: result.applied.length,
    issues: result.issues.map((issue) => JSON.stringify(issue)),
    bytes: (
      await serializeCanonicalSave({ snapshot: session.captureSaveSnapshot() })
    ).buffer,
  };
};

const viaDefaultEngine = async (source: ArrayBuffer, mode: Mode) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source.slice(0));
  reviewer.applyDocumentOperations(
    { version: 1, mode, operations: [EDIT] },
    { undefinedReferences: "refuse" },
  );
  return reviewer.toBuffer();
};

describe.each(Object.entries(ROWS))(
  "in the canonical session, a table with a row of %s",
  (_name, row) => {
    test.each(["direct", "tracked-changes"] as const)(
      "saves the row as the file states it after an edit beside it (%s)",
      async (mode) => {
        const source = await docx(row);
        const canonical = await viaCanonicalEngine(source, mode);
        expect(canonical.issues).toEqual([]);
        expect(canonical.applied).toBe(1);
        const rows = await savedRows(canonical.bytes);
        const sourceRows = await savedRows(source);
        expect(rows).toHaveLength(3);
        // The middle row keeps its cells and omissions exactly.
        expect(rows[1]).toEqual(sourceRows[1]);
        expect(rows).toEqual(
          await savedRows(await viaDefaultEngine(source, mode)),
        );
      },
    );
  },
);
