import { panic } from "better-result";
import { expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  buildRawPackage,
  SECTION,
  WORD_STYLES,
} from "../__tests__/documentShapes";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(" ");
const BOX =
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1828800" cy="457200"/><wp:docPr id="2" name="Box"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  "<wps:txbx><w:txbxContent><w:p><w:r><w:t>Boxed words</w:t></w:r></w:p></w:txbxContent></wps:txbx>" +
  "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

const NOTE_KINDS = {
  footnote: { part: "word/footnotes.xml", root: "footnotes" },
  endnote: { part: "word/endnotes.xml", root: "endnotes" },
} as const;
type NoteKind = keyof typeof NOTE_KINDS;

/** A body paragraph referencing note 1, whose paragraph draws a text box. */
const notePackage = async (kind: NoteKind) => {
  const { part, root } = NOTE_KINDS[kind];
  const bytes = await buildRawPackage({
    styles: WORD_STYLES,
    body: `<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:${kind}Reference w:id="1"/></w:r></w:p>${SECTION()}`,
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
          `<w:${kind} w:id="1"><w:p><w:r><w:t>Note</w:t></w:r>${BOX}</w:p></w:${kind}>` +
          `</w:${root}>`,
      },
    ],
  });
  return Uint8Array.from(bytes).buffer;
};

/** The text the saved text box shows. */
const savedBoxText = (xml: string): string =>
  [
    ...xml
      .slice(xml.indexOf("<w:txbxContent>"), xml.indexOf("</w:txbxContent>"))
      .matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu),
  ]
    .map(([, words]) => words)
    .join("");

// A save that writes a note again writes the text box the note draws: the
// parser reads the box as it reads one in the body, for every engine.
for (const session of ["canonical"] as const)
  for (const kind of ["footnote", "endnote"] as const)
    test(`${session}: an edit to a ${kind}'s own text keeps the text box it draws`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await notePackage(kind),
        { session },
      );
      const notes =
        kind === "footnote"
          ? reviewer.toDocument().package.footnotes
          : reviewer.toDocument().package.endnotes;
      const paragraph = notes?.find(({ id }) => id === 1)?.content.at(0);
      const blockId =
        paragraph?.type === "paragraph"
          ? (paragraph.paraId ?? panic("The note paragraph holds no id."))
          : panic("No note paragraph");
      const result = reviewer.applyDocumentOperationsToStory({
        story: { type: kind, noteId: 1 },
        batch: {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "direct",
          operations: [
            {
              id: "edit",
              type: "replaceInBlock",
              blockId,
              find: "Note",
              replace: "Remark",
            },
          ],
        },
        undefinedReferences: "refuse",
      });
      expect(result.skipped).toEqual([]);
      const xml =
        (await (
          await ZipArchive.read(await reviewer.toBuffer())
        )
          .entry(NOTE_KINDS[kind].part)
          ?.text()) ?? "";
      expect(xml).toContain("Remark");
      expect(xml.split("<w:txbxContent>")).toHaveLength(2);
      expect(savedBoxText(xml)).toBe("Boxed words");
    });
