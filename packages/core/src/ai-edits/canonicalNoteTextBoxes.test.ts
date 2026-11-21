import { panic } from "better-result";
import { afterAll, beforeAll, expect, setSystemTime, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  buildRawPackage,
  SECTION,
  WORD_STYLES,
} from "../__tests__/documentShapes";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationMode,
} from "../document-operations";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

beforeAll(() => setSystemTime(new Date("2026-01-01T12:00:00Z")));
afterAll(() => setSystemTime());

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(" ");
const box = (text: string) =>
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1828800" cy="457200"/><wp:docPr id="2" name="Box"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
  "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

const NOTE_KINDS = {
  footnote: {
    part: "word/footnotes.xml",
    root: "footnotes",
    reference: "footnoteReference",
  },
  endnote: {
    part: "word/endnotes.xml",
    root: "endnotes",
    reference: "endnoteReference",
  },
} as const;
type NoteKind = keyof typeof NOTE_KINDS;

/** A body paragraph referencing note 1, whose paragraph draws a text box. */
const notePackage = async (kind: NoteKind) => {
  const { part, root, reference } = NOTE_KINDS[kind];
  const bytes = await buildRawPackage({
    styles: WORD_STYLES,
    body: `<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:${reference} w:id="1"/></w:r></w:p>${SECTION()}`,
    parts: [
      {
        path: part,
        contentType: `application/vnd.openxmlformats-officedocument.wordprocessingml.${root}+xml`,
        relationshipType: root,
        relationshipId: `rId${root}`,
        body:
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${root} ${NAMESPACES}>` +
          `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
          `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>` +
          `<w:${kind} w:id="1"><w:p><w:r><w:t>Note</w:t></w:r>${box("Boxed words")}</w:p></w:${kind}>` +
          `</w:${root}>`,
      },
    ],
  });
  return Uint8Array.from(bytes).buffer;
};

const open = (buffer: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(buffer, { session: "canonical" });

const noteOf = (document: Document, kind: NoteKind) =>
  (kind === "footnote"
    ? document.package.footnotes
    : document.package.endnotes
  )?.find(({ id }) => id === 1) ?? panic(`No ${kind} 1`);

/** The paragraphs of the text box the note's paragraph draws. */
const boxParagraphs = (document: Document, kind: NoteKind) => {
  const host = noteOf(document, kind).content.at(0);
  if (host?.type !== "paragraph") return panic("No note paragraph");
  for (const run of host.content)
    if (run.type === "run")
      for (const item of run.content)
        if (item.type === "shape" && item.shape.textBody !== undefined)
          return item.shape.textBody.content.filter(
            (block) => block.type === "paragraph",
          );
  return panic("The note no longer draws the box.");
};

/** The text a reader keeps of the box: deleted text left out. */
const boxText = (document: Document, kind: NoteKind): string =>
  JSON.stringify(boxParagraphs(document, kind))
    .match(/"text":"[^"]*"/gu)
    ?.map((text) => text.slice(8, -1))
    .join("") ?? "";

const savedPart = async (reviewer: FolioDocxReviewer, kind: NoteKind) =>
  (await (
    await ZipArchive.read(await reviewer.toBuffer())
  )
    .entry(NOTE_KINDS[kind].part)
    ?.text()) ?? "";

/** The saved text box's markup, and the text it shows (deleted text left out). */
const savedBox = (xml: string) => {
  const markup = xml.slice(
    xml.indexOf("<w:txbxContent>"),
    xml.indexOf("</w:txbxContent>"),
  );
  const text = [...markup.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)]
    .map(([, words]) => words)
    .join("");
  return { markup, text };
};

const edit = (
  reviewer: FolioDocxReviewer,
  kind: NoteKind,
  mode: FolioDocumentOperationMode,
) => {
  const boxId =
    boxParagraphs(reviewer.toDocument(), kind).at(0)?.paraId ??
    panic("The box paragraph holds no id.");
  return reviewer.applyDocumentOperationsToStory({
    story: { type: kind, noteId: 1 },
    batch: {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [
        {
          id: "edit",
          type: "replaceInBlock",
          blockId: boxId,
          find: "Boxed",
          replace: "Framed",
        },
      ],
    },
    undefinedReferences: "refuse",
  });
};

for (const kind of ["footnote", "endnote"] as const) {
  test(`a ${kind}'s text box paragraph opens with an id that survives a save and a reopen`, async () => {
    const reviewer = await open(await notePackage(kind));
    const [paragraph] = boxParagraphs(reviewer.toDocument(), kind);
    expect(paragraph?.paraId).toMatch(/^[0-9A-F]{8}$/u);
    const reopened = await open(await reviewer.toBuffer());
    expect(boxParagraphs(reopened.toDocument(), kind).at(0)?.paraId).toBe(
      paragraph?.paraId,
    );
  });

  test(`a direct edit in a ${kind}'s text box saves into the box and reopens`, async () => {
    const reviewer = await open(await notePackage(kind));
    const before = boxParagraphs(reviewer.toDocument(), kind).at(0)?.paraId;
    const result = edit(reviewer, kind, "direct");
    expect(result.skipped).toEqual([]);
    expect(result.applied.map(({ id }) => id)).toEqual(["edit"]);
    const xml = await savedPart(reviewer, kind);
    expect(xml.split("<w:txbxContent>")).toHaveLength(2);
    expect(savedBox(xml).text).toBe("Framed words");
    const reopened = await open(await reviewer.toBuffer());
    expect(boxText(reopened.toDocument(), kind)).toBe("Framed words");
    expect(boxParagraphs(reopened.toDocument(), kind).at(0)?.paraId).toBe(
      before,
    );
    expect(noteOf(reopened.toDocument(), kind).content.at(0)).toMatchObject({
      type: "paragraph",
    });
  });

  for (const decision of ["accept", "reject"] as const)
    test(`a tracked edit in a ${kind}'s text box saves as a revision, reopens and is ${decision}ed`, async () => {
      const reviewer = await open(await notePackage(kind));
      const result = edit(reviewer, kind, "tracked-changes");
      expect(result.skipped).toEqual([]);
      const { markup } = savedBox(await savedPart(reviewer, kind));
      expect(markup).toContain("<w:ins ");
      expect(markup).toContain("<w:del ");
      const reopened = await open(await reviewer.toBuffer());
      const resolution =
        decision === "accept" ? reopened.acceptAll() : reopened.rejectAll();
      expect(resolution.resolved).toBeGreaterThan(0);
      const resolved = await open(await reopened.toBuffer());
      const expected = decision === "accept" ? "Framed words" : "Boxed words";
      expect(boxText(resolved.toDocument(), kind)).toBe(expected);
      const saved = savedBox(await savedPart(resolved, kind));
      expect(saved.text).toBe(expected);
      expect(saved.markup).not.toContain("<w:ins ");
      expect(saved.markup).not.toContain("<w:del ");
    });
}
