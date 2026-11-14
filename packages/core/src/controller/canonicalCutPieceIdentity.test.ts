/**
 * Replacing the first words of another reviewer's pending insertion cuts it:
 * the struck words stay in the first piece, the new words follow, and the
 * rest of the insertion is a second piece. Accepting the replacement leaves
 * the insertion with its own id, as the same edit made directly does, also
 * after the file is saved and opened again.
 */
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  createCanonicalEditorHarness,
  saveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { parseDocx } from "../docx/parser";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import type { Document } from "../types/document";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const source = async () => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      `<w:p w14:paraId="11111111"><w:r><w:t xml:space="preserve">Lead </w:t></w:r>` +
      `<w:ins w:id="7" w:author="Earlier" w:date="2024-01-01T00:00:00Z"><w:r><w:t>first words</w:t></w:r></w:ins></w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const parse = (bytes: ArrayBuffer) =>
  parseDocx(bytes, { preloadFonts: false, detectVariables: false });

type Revision = { id: number; author: string; type: string; text: string };

/** The insertions and deletions of the main story, in order. */
const revisions = (document: Document): Revision[] => {
  const out: Revision[] = [];
  const text = (nodes: readonly unknown[]): string =>
    nodes
      .map((node) => {
        const record = node as {
          type: string;
          text?: string;
          content?: unknown[];
        };
        if (record.type === "text") return record.text ?? "";
        return text(record.content ?? []);
      })
      .join("");
  const visit = (nodes: readonly unknown[]) => {
    for (const node of nodes) {
      const record = node as {
        type: string;
        info?: { id: number; author: string };
        content?: unknown[];
      };
      if (
        (record.type === "insertion" || record.type === "deletion") &&
        record.info !== undefined
      )
        out.push({
          id: record.info.id,
          author: record.info.author,
          type: record.type,
          text: text(record.content ?? []),
        });
      visit(record.content ?? []);
    }
  };
  for (const block of document.package.document.content)
    if (block.type === "paragraph") visit(block.content);
  return out;
};

/** Accept one reviewer's changes in a fresh session over `document`. */
const acceptBy = (document: Document, author: string) => {
  const session = createCanonicalSession(document).unwrap();
  const state = EditorState.create({ doc: session.projection.doc });
  const ids = revisions(session.document)
    .filter((revision) => revision.author === author)
    .map(({ id }) => id);
  const commit = session
    .prepareResolve(state, { revisionIds: ids, resolution: "accept" })
    .unwrap();
  publishCanonicalProjection({ state, session, commit }).unwrap();
  return session.captureSaveSnapshot().document;
};

const typedOver = async (mode: "editing" | "suggesting") => {
  const harness = createCanonicalEditorHarness(
    await parse(await source()),
    mode,
  );
  try {
    let from = 0;
    harness.state.doc.descendants((node, position) => {
      const index = node.isText ? (node.text ?? "").indexOf("first") : -1;
      if (index !== -1 && from === 0) from = position + index;
      return true;
    });
    harness.dispatch(
      harness.state.tr.setSelection(
        TextSelection.create(harness.state.doc, from, from + 5),
      ),
    );
    harness.typeText("New");
    expect(harness.refusals).toEqual([]);
    return await saveCanonicalHarnessDocument(harness.snapshot());
  } finally {
    harness.dispose();
  }
};

describe("replacing the start of another reviewer's insertion", () => {
  test("keeps that insertion's id once the replacement is accepted, after save and reopen", async () => {
    const direct = revisions(
      await parse((await typedOver("editing")).bytes.slice().buffer),
    );
    expect(direct).toEqual([
      { id: 7, author: "Earlier", type: "insertion", text: " words" },
    ]);
    const saved = await typedOver("suggesting");
    const xml =
      (await (
        await ZipArchive.read(saved.bytes)
      )
        .entry("word/document.xml")
        ?.text()) ?? "";
    expect(xml).toContain("folio:retainedEnclosing=");
    const pending = revisions(saved.model);
    expect(
      pending.map(({ author, type, text }) => ({ author, type, text })),
    ).toEqual([
      { author: "Earlier", type: "insertion", text: "first" },
      { author: "Conformance", type: "deletion", text: "first" },
      { author: "Conformance", type: "insertion", text: "New" },
      { author: "Earlier", type: "insertion", text: " words" },
    ]);
    for (const document of [
      saved.model,
      await parse(saved.bytes.slice().buffer),
    ]) {
      const accepted = revisions(acceptBy(document, "Conformance"));
      expect(accepted).toEqual(direct);
    }
  });
});
