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
import { markupViewNotes } from "../prosemirror/markupViewNotes";
import { createHarnessState, parseShapeDocument } from "./editorHarness";

type Fixture = { body: string; footnotes?: string; header?: string };
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

const packageOf = async ({ body, footnotes, header }: Fixture): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${MAIN}.styles+xml"/><Override PartName="/word/settings.xml" ContentType="${MAIN}.settings+xml"/><Override PartName="/word/footnotes.xml" ContentType="${MAIN}.footnotes+xml"/>${header === undefined ? "" : `<Override PartName="/word/header1.xml" ContentType="${MAIN}.header+xml"/>`}</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIPS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIPS}/styles" Target="styles.xml"/><Relationship Id="rId3" Type="${RELATIONSHIPS}/settings" Target="settings.xml"/><Relationship Id="rId4" Type="${RELATIONSHIPS}/footnotes" Target="footnotes.xml"/>${header === undefined ? "" : `<Relationship Id="rId5" Type="${RELATIONSHIPS}/header" Target="header1.xml"/>`}</Relationships>`,
  );
  const sectioned = body.includes("<w:sectPr") ? body : `${body}${reference.package.sectPr}`;
  const withSection =
    header === undefined
      ? sectioned
      : sectioned.replace(
          "<w:sectPr>",
          '<w:sectPr><w:headerReference w:type="default" r:id="rId5"/>',
        );
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
  if (header !== undefined) {
    zip.file(
      "word/header1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NAMESPACES}>${header}</w:hdr>`,
    );
  }
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
    expect(noteState(restoreNotes(document, restored, state.doc), 1)).toEqual({
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

  test("the markup views show the notes as they resolve the body", async () => {
    const document = await open(fixture);
    const state = createHarnessState(document, "editing");
    const shown = (view: "all-markup" | "original" | "no-markup") => {
      const footnotes = markupViewNotes(document.package.footnotes, view, state, document);
      return noteState(
        { ...document, package: { ...document.package, ...(footnotes && { footnotes }) } },
        1,
      );
    };
    expect(shown("all-markup")).toEqual({ text: "The note text.", deleted: true });
    expect(shown("original")).toEqual({ text: "The note text.", deleted: false });
    expect(shown("no-markup")).toEqual({ text: "", deleted: false });
  });

  test("a save right after a restore reads the restored note", async () => {
    const document = await open(fixture);
    const deleted = createHarnessState(document, "editing");
    const rejected = resolve(deleted, "reject");
    const follower = createNoteReferenceFollower();
    follower.noteBase(deleted.doc);
    follower.reconcile(document, rejected.doc);
    // The host has not handed the written-back document to the editor yet.
    expect(noteState(follower.withPending(document), 1)).toEqual({
      text: "The note text.",
      deleted: false,
    });
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

describe("a note shared with other revisions and other stories", () => {
  const fixture = reference.fixtures["authored-delete-footnote-reference"];
  if (!fixture?.footnotes) throw new Error("missing the deleted reference fixture");
  // Deleted earlier, by someone else: not part of the reference's deletion.
  const struckEarlier =
    '<w:del w:id="3" w:author="Other" w:date="2025-06-01T00:00:00Z"><w:r><w:delText xml:space="preserve"> Struck earlier.</w:delText></w:r></w:del></w:p></w:footnote>';
  const withEarlierDeletion: Fixture = {
    ...fixture,
    footnotes: fixture.footnotes.replace("</w:p></w:footnote>", struckEarlier),
  };

  /** The note's text still pending deletion. */
  const deletedText = (document: Document, id: number): string => {
    const note = document.package.footnotes?.find((candidate) => candidate.id === id);
    let text = "";
    footnoteToProseDoc(note?.content ?? []).descendants((node) => {
      if (node.isText && node.marks.some((mark) => mark.type.name === "deletion")) {
        text += node.text ?? "";
      }
      return true;
    });
    return text.trim();
  };

  test("the reviewer rejecting the reference's deletion keeps the note's other deletion", async () => {
    const bytes = await packageOf(withEarlierDeletion);
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer);
    const state = createHarnessState(await open(withEarlierDeletion), "editing");
    expect(reviewer.rejectChange(referenceDeletion(state.doc))).toBe(true);
    const saved = await parseShapeDocument(new Uint8Array(await reviewer.toBuffer()));
    expect(noteState(saved, 1)?.text).toBe("The note text. Struck earlier.");
    expect(deletedText(saved, 1)).toBe("Struck earlier.");
  });

  test("an editor rejecting the reference's deletion keeps the note's other deletion", async () => {
    const document = await open(withEarlierDeletion);
    const deleted = createHarnessState(document, "editing");
    const rejected = resolve(deleted, "reject");
    const follower = createNoteReferenceFollower();
    follower.noteBase(deleted.doc);
    const restored = follower.reconcile(document, rejected.doc);
    expect(noteState(restored, 1)?.text).toBe("The note text. Struck earlier.");
    expect(deletedText(restored, 1)).toBe("Struck earlier.");
    const keys = noteReferencesRestored(deleted.doc, rejected.doc);
    expect(deletedText(restoreNotes(document, keys, deleted.doc), 1)).toBe("Struck earlier.");
  });

  test("accepting the body reference's deletion keeps a note a header still refers to", async () => {
    const withHeader: Fixture = {
      ...fixture,
      header:
        '<w:p w14:paraId="7A000007"><w:r><w:t>Header</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>',
    };
    const bytes = await packageOf(withHeader);
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer);
    const state = createHarnessState(await open(withHeader), "editing");
    expect(reviewer.acceptChange(referenceDeletion(state.doc))).toBe(true);
    expect(reviewer.listStories().some(({ handle }) => handle.type === "footnote")).toBe(true);
    const saved = await parseShapeDocument(new Uint8Array(await reviewer.toBuffer()));
    expect(noteState(saved, 1)?.text).toBe("The note text.");
  });
});

describe("an edit operation that deletes a note reference", () => {
  const deleted = reference.fixtures["authored-delete-footnote-reference"];
  const liveNote = reference.fixtures["nt-reference-only-deleted"]?.footnotes;
  if (!deleted || liveNote === undefined) throw new Error("missing the note reference fixtures");
  // The same body with its reference live, and a note whose text is not deleted.
  const live: Fixture = {
    body: deleted.body.replace(
      /<w:del\b[^>]*>(<w:r>.*?<w:footnoteReference w:id="1"\/><\/w:r>)<\/w:del>/u,
      "$1",
    ),
    footnotes: liveNote,
  };

  const deleteNotedParagraph = async (mode: "direct" | "tracked-changes") => {
    const reviewer = await FolioDocxReviewer.fromBuffer((await packageOf(live)).slice().buffer);
    const block = reviewer.getContent().find(({ text }) => text.startsWith("Noted text"));
    if (!block) throw new Error("no noted paragraph");
    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode,
      operations: [{ id: "delete", type: "deleteBlock", blockId: block.id }],
    });
    expect(result.skipped).toEqual([]);
    return { reviewer, undoHandle: result.undoHandle };
  };
  const saved = async (reviewer: FolioDocxReviewer): Promise<Document> =>
    parseShapeDocument(new Uint8Array(await reviewer.toBuffer()));

  test("the fixture's reference and note are live", async () => {
    expect(live.body).not.toContain("<w:del ");
    expect(noteState(await open(live), 1)).toEqual({ text: "The note text.", deleted: false });
  });

  test("tracked, it deletes the note's text with the reference", async () => {
    const { reviewer } = await deleteNotedParagraph("tracked-changes");
    expect(noteState(await saved(reviewer), 1)).toEqual({ text: "The note text.", deleted: true });
  });

  test("tracked and accepted, the note goes", async () => {
    const { reviewer } = await deleteNotedParagraph("tracked-changes");
    expect(reviewer.acceptAll()).toBeGreaterThan(0);
    expect(noteState(await saved(reviewer), 1)).toBeNull();
  });

  test("tracked and rejected, the note keeps its text", async () => {
    const { reviewer } = await deleteNotedParagraph("tracked-changes");
    expect(reviewer.rejectAll()).toBeGreaterThan(0);
    expect(noteState(await saved(reviewer), 1)).toEqual({ text: "The note text.", deleted: false });
  });

  test("tracked and undone, the note keeps its text", async () => {
    const { reviewer, undoHandle } = await deleteNotedParagraph("tracked-changes");
    if (!undoHandle) throw new Error("no undo handle");
    expect(reviewer.undoDocumentOperations(undoHandle).status).toBe("undone");
    expect(noteState(await saved(reviewer), 1)).toEqual({ text: "The note text.", deleted: false });
  });

  test("direct, the note goes with its reference", async () => {
    const { reviewer } = await deleteNotedParagraph("direct");
    expect(noteState(await saved(reviewer), 1)).toBeNull();
  });
});
