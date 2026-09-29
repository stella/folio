/**
 * A note goes with its reference: resolving a reference's deletion in the
 * body resolves the note's story to match, as the reference implementation
 * does.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import JSZip from "jszip";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";

import {
  createNoteReferenceFollower,
  noteReferencesRestored,
  restoreNotes,
  withoutUnreferencedNotes,
} from "../prosemirror/noteReferenceReview";
import { acceptAIEditRevision, rejectAIEditRevision } from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { footnoteToProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createHarnessState, parseShapeDocument } from "./editorHarness";

type Fixture = { body: string; footnotes?: string };
const reference = JSON.parse(
  readFileSync(
    path.join(import.meta.dir, "__fixtures__", "revision-resolution-reference.json"),
    "utf8",
  ),
) as {
  package: { styles: string; numbering: string; settings: string; sectPr: string };
  fixtures: Record<string, Fixture>;
};

const MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NAMESPACES =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  `xmlns:r="${RELATIONSHIPS}" ` +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"';
const SEPARATORS =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';

const packageOf = async ({ body, footnotes }: Fixture): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${MAIN}.styles+xml"/><Override PartName="/word/settings.xml" ContentType="${MAIN}.settings+xml"/><Override PartName="/word/footnotes.xml" ContentType="${MAIN}.footnotes+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIPS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIPS}/styles" Target="styles.xml"/><Relationship Id="rId3" Type="${RELATIONSHIPS}/settings" Target="settings.xml"/><Relationship Id="rId4" Type="${RELATIONSHIPS}/footnotes" Target="footnotes.xml"/></Relationships>`,
  );
  const withSection = body.includes("<w:sectPr") ? body : `${body}${reference.package.sectPr}`;
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NAMESPACES}><w:body>${withSection}</w:body></w:document>`,
  );
  zip.file("word/styles.xml", reference.package.styles);
  zip.file("word/settings.xml", reference.package.settings);
  zip.file(
    "word/footnotes.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${NAMESPACES}>${SEPARATORS}${footnotes ?? ""}</w:footnotes>`,
  );
  return zip.generateAsync({ type: "uint8array" });
};

const open = async (fixture: Fixture): Promise<Document> =>
  parseShapeDocument(await packageOf(fixture));

/** The revision that deletes the body's note reference. */
const referenceDeletion = (doc: PMNode): number => {
  let id: number | null = null;
  doc.descendants((node) => {
    if (node.marks.some((mark) => mark.type.name === "footnoteRef")) {
      const deletion = node.marks.find((mark) => mark.type.name === "deletion");
      if (deletion) id = deletion.attrs["revisionId"] as number;
    }
    return id === null;
  });
  if (id === null) throw new Error("no deleted note reference");
  return id;
};

/** A note's text, and whether any of it is still pending deletion. */
const noteState = (document: Document, id: number) => {
  const note = document.package.footnotes?.find((candidate) => candidate.id === id);
  if (!note) return null;
  const doc = footnoteToProseDoc(note.content);
  let deleted = false;
  doc.descendants((node) => {
    if (node.marks.some((mark) => mark.type.name === "deletion")) deleted = true;
    if ((node.attrs["pPrMark"] as { kind?: string } | null)?.kind === "del") deleted = true;
    return true;
  });
  return { text: doc.textContent.trim(), deleted };
};

const resolve = (state: EditorState, mode: "accept" | "reject"): EditorState => {
  const command = mode === "accept" ? acceptAIEditRevision : rejectAIEditRevision;
  let next = state;
  command(referenceDeletion(state.doc))(state, (transaction) => {
    next = state.apply(transaction);
  });
  return next;
};

describe("a note goes with its reference", () => {
  const fixture = reference.fixtures["authored-delete-footnote-reference"];
  if (!fixture) throw new Error("missing the deleted reference fixture");

  test("rejecting the reference's deletion gives the note its text back", async () => {
    const document = await open(fixture);
    expect(noteState(document, 1)).toEqual({ text: "The note text.", deleted: true });
    const state = createHarnessState(document, "editing");
    const rejected = resolve(state, "reject");
    const restored = noteReferencesRestored(state.doc, rejected.doc);
    expect(restored).toEqual(["footnote:1"]);
    expect(noteState(restoreNotes(document, restored), 1)).toEqual({
      text: "The note text.",
      deleted: false,
    });
  });

  test("accepting the reference's deletion drops the note on save", async () => {
    const document = await open(fixture);
    const state = createHarnessState(document, "editing");
    const accepted = resolve(state, "accept");
    expect(noteReferencesRestored(state.doc, accepted.doc)).toEqual([]);
    const model = fromProseDoc(accepted.doc, document);
    const saved = withoutUnreferencedNotes(model);
    expect(noteState(saved, 1)).toBeNull();
    expect(saved.package.footnotes?.length).toBe((model.package.footnotes?.length ?? 0) - 1);
    expect(withoutUnreferencedNotes(fromProseDoc(state.doc, document))).toEqual(
      fromProseDoc(state.doc, document),
    );
  });

  test("a note dropped on save takes the comments anchored only in it", async () => {
    const document = await open(fixture);
    const accepted = resolve(createHarnessState(document, "editing"), "accept");
    const model = fromProseDoc(accepted.doc, document);
    const note = model.package.footnotes?.find(({ id }) => id === 1);
    const first = note?.content[0];
    if (first?.type !== "paragraph") throw new Error("missing the note's paragraph");
    first.content.unshift({ type: "commentRangeStart", id: 7 });
    first.content.push({ type: "commentRangeEnd", id: 7 });
    model.package.document.comments = [
      { id: 7, author: "Reviewer", date: "2026-01-01T00:00:00Z", content: [] },
    ];
    expect(withoutUnreferencedNotes(model).package.document.comments).toEqual([]);
  });

  test("an editor's notes follow the reject and its undo", async () => {
    const document = await open(fixture);
    const deleted = createHarnessState(document, "editing");
    const rejected = resolve(deleted, "reject");
    const follower = createNoteReferenceFollower();
    follower.noteBase(deleted.doc);
    const restored = follower.reconcile(document, rejected.doc);
    expect(noteState(restored, 1)).toEqual({ text: "The note text.", deleted: false });
    // Undo puts the reference's deletion back: the note returns as it was.
    follower.noteBase(rejected.doc);
    const undone = follower.reconcile(restored, deleted.doc);
    expect(undone.package.footnotes).toEqual(document.package.footnotes);
  });

  test("a reference deleted again after its note changed takes the note's text", async () => {
    const document = await open(fixture);
    const deleted = createHarnessState(document, "editing");
    const rejected = resolve(deleted, "reject");
    const follower = createNoteReferenceFollower();
    follower.noteBase(deleted.doc);
    const restored = follower.reconcile(document, rejected.doc);
    const edited: Document = {
      ...restored,
      package: {
        ...restored.package,
        footnotes: restored.package.footnotes?.map((note) => Object.assign({}, note)),
      },
    };
    follower.noteBase(rejected.doc);
    const redeleted = follower.reconcile(edited, deleted.doc);
    expect(noteState(redeleted, 1)).toEqual({ text: "The note text.", deleted: true });
  });

  for (const mode of ["accept", "reject"] as const) {
    test(`the reviewer resolving only the reference's deletion (${mode}) resolves the note too`, async () => {
      const bytes = await packageOf(fixture);
      const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer);
      const state = createHarnessState(await open(fixture), "editing");
      const id = referenceDeletion(state.doc);
      expect(mode === "accept" ? reviewer.acceptChange(id) : reviewer.rejectChange(id)).toBe(true);
      const saved = await parseShapeDocument(new Uint8Array(await reviewer.toBuffer()));
      expect(noteState(saved, 1)).toEqual(
        mode === "accept" ? null : { text: "The note text.", deleted: false },
      );
    });
  }
});
